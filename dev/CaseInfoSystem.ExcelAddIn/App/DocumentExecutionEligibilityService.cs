using System;
using System.Collections.Generic;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentExecutionEligibilityService
    {
        private const string TaskPaneMasterVersionPropertyName = "TASKPANE_MASTER_VERSION";

        private readonly ExcelInteropService _excelInteropService;
        private readonly DocumentTemplateResolver _documentTemplateResolver;
        private readonly CaseContextFactory _caseContextFactory;
        private readonly DocumentOutputService _documentOutputService;
        private readonly Logger _logger;
        private readonly Dictionary<string, DocumentExecutionEligibility> _eligibleCacheByWorkbookKey;

        /// <summary>
        internal DocumentExecutionEligibilityService(
            ExcelInteropService excelInteropService,
            DocumentTemplateResolver documentTemplateResolver,
            CaseContextFactory caseContextFactory,
            DocumentOutputService documentOutputService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _documentTemplateResolver = documentTemplateResolver ?? throw new ArgumentNullException(nameof(documentTemplateResolver));
            _caseContextFactory = caseContextFactory ?? throw new ArgumentNullException(nameof(caseContextFactory));
            _documentOutputService = documentOutputService ?? throw new ArgumentNullException(nameof(documentOutputService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _eligibleCacheByWorkbookKey = new Dictionary<string, DocumentExecutionEligibility>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        internal DocumentExecutionEligibility Evaluate(Excel.Workbook workbook, string actionKind, string key)
        {
            Stopwatch totalStopwatch = Stopwatch.StartNew();
            Stopwatch phaseStopwatch = Stopwatch.StartNew();
            if (workbook == null)
            {
                return new DocumentExecutionEligibility(false, "workbook was null", null);
            }

            if (!string.Equals(actionKind, "doc", StringComparison.OrdinalIgnoreCase))
            {
                return new DocumentExecutionEligibility(false, "action was not doc", null);
            }

            string normalizedKey = (key ?? string.Empty).Trim();
            if (normalizedKey.Length == 0)
            {
                return new DocumentExecutionEligibility(false, "document key was empty", null);
            }

            string cacheKey = BuildEligibleCacheKey(workbook, actionKind, normalizedKey);
            if (cacheKey.Length > 0
                && _eligibleCacheByWorkbookKey.TryGetValue(cacheKey, out DocumentExecutionEligibility cachedEligibility))
            {
                _logger.Debug(
                    "DocumentExecutionEligibilityService.Evaluate",
                    "CacheHit elapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed)
                    + " key=" + normalizedKey
                    + " canExecute=" + cachedEligibility.CanExecuteInVsto.ToString());
                return new DocumentExecutionEligibility(cachedEligibility.CanExecuteInVsto, cachedEligibility.Reason, cachedEligibility.TemplateSpec, null);
            }

            DocumentTemplateSpec templateSpec = _documentTemplateResolver.Resolve(workbook, normalizedKey);
            _logger.Debug(
                "DocumentExecutionEligibilityService.Evaluate",
                "TemplateResolved elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed)
                + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed)
                + " key=" + normalizedKey
                + " resolved=" + (templateSpec != null).ToString());
            phaseStopwatch.Restart();
            if (templateSpec == null)
            {
                return new DocumentExecutionEligibility(false, "template spec was not resolved", null);
            }

            if (!DocumentTemplateResolver.IsSupportedWordTemplate(templateSpec))
            {
                return new DocumentExecutionEligibility(false, "template type is not supported: " + (templateSpec.TemplatePath ?? string.Empty), templateSpec);
            }

            if (IsMacroEnabledTemplate(templateSpec))
            {
                return new DocumentExecutionEligibility(false, "macro-enabled word template is routed to VBA: " + (templateSpec.TemplatePath ?? string.Empty), templateSpec);
            }

            if (!_documentTemplateResolver.TemplateExists(templateSpec))
            {
                return new DocumentExecutionEligibility(false, "template file was not found: " + (templateSpec.TemplatePath ?? string.Empty), templateSpec);
            }

            string outputFolder = _documentOutputService.ResolveWorkbookFolder(workbook);
            _logger.Debug(
                "DocumentExecutionEligibilityService.Evaluate",
                "OutputFolderResolved elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed)
                + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed)
                + " key=" + normalizedKey
                + " resolved=" + (!string.IsNullOrWhiteSpace(outputFolder)).ToString());
            phaseStopwatch.Restart();
            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                return new DocumentExecutionEligibility(false, "output folder could not be resolved", templateSpec);
            }

            CaseContext caseContext = _caseContextFactory.CreateForDocumentCreate(workbook);
            _logger.Debug(
                "DocumentExecutionEligibilityService.Evaluate",
                "CaseContextResolved elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed)
                + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed)
                + " key=" + normalizedKey
                + " resolved=" + (caseContext != null).ToString());
            if (caseContext == null)
            {
                return new DocumentExecutionEligibility(false, "case context could not be resolved", templateSpec);
            }

            if (caseContext.CaseValues == null || caseContext.CaseValues.Count == 0)
            {
                return new DocumentExecutionEligibility(false, "case snapshot could not be resolved", templateSpec);
            }

            if (string.IsNullOrWhiteSpace(templateSpec.DocumentName))
            {
                _logger.Warn("DocumentExecutionEligibilityService found empty document name. key=" + normalizedKey);
            }

            _logger.Info(
                "DocumentExecutionEligibilityService marked template eligible."
                + " key=" + normalizedKey
                + ", source=" + templateSpec.ResolutionSource
                + ", template=" + (templateSpec.TemplatePath ?? string.Empty)
                + ", outputFolder=" + outputFolder
                + ", elapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed));

            var eligibility = new DocumentExecutionEligibility(true, "eligible", templateSpec, caseContext);
            if (cacheKey.Length > 0)
            {
                _eligibleCacheByWorkbookKey[cacheKey] = new DocumentExecutionEligibility(true, "eligible", templateSpec, null);
            }

            return eligibility;
        }

        private string BuildEligibleCacheKey(Excel.Workbook workbook, string actionKind, string normalizedKey)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            string workbookFullName;
            try
            {
                workbookFullName = workbook.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }

            if (workbookFullName.Length == 0)
            {
                return string.Empty;
            }

            string masterVersion = _excelInteropService.TryGetDocumentProperty(workbook, TaskPaneMasterVersionPropertyName) ?? string.Empty;
            return workbookFullName + "|" + masterVersion + "|" + (actionKind ?? string.Empty) + "|" + normalizedKey;
        }

        /// <summary>
        private static bool IsMacroEnabledTemplate(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null || string.IsNullOrWhiteSpace(templateSpec.TemplatePath))
            {
                return false;
            }

            string extension = System.IO.Path.GetExtension(templateSpec.TemplatePath) ?? string.Empty;
            return string.Equals(extension, ".dotm", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".docm", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatElapsedSeconds(TimeSpan elapsed)
        {
            return elapsed.TotalSeconds.ToString("0.000");
        }
    }
}
