using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentExecutionEligibilityService
    {
        private readonly DocumentTemplateResolver _documentTemplateResolver;
        private readonly CaseContextFactory _caseContextFactory;
        private readonly DocumentOutputService _documentOutputService;
        private readonly Logger _logger;

        /// <summary>
        internal DocumentExecutionEligibilityService(
            DocumentTemplateResolver documentTemplateResolver,
            CaseContextFactory caseContextFactory,
            DocumentOutputService documentOutputService,
            Logger logger)
        {
            _documentTemplateResolver = documentTemplateResolver ?? throw new ArgumentNullException(nameof(documentTemplateResolver));
            _caseContextFactory = caseContextFactory ?? throw new ArgumentNullException(nameof(caseContextFactory));
            _documentOutputService = documentOutputService ?? throw new ArgumentNullException(nameof(documentOutputService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
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
                + ", " + BuildPathDiagnostics("template", templateSpec.TemplatePath)
                + ", " + BuildFolderDiagnostics("outputFolder", outputFolder)
                + ", elapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed));

            return new DocumentExecutionEligibility(true, "eligible", templateSpec, caseContext);
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

        private static string BuildPathDiagnostics(string label, string path)
        {
            string safeLabel = label ?? string.Empty;
            string safePath = path ?? string.Empty;
            return safeLabel + "Present=" + (!string.IsNullOrWhiteSpace(safePath)).ToString()
                + ", " + safeLabel + "Length=" + safePath.Length.ToString()
                + ", " + safeLabel + "Extension=" + SafeGetExtension(safePath)
                + ", " + safeLabel + "Exists=" + SafeFileExists(safePath).ToString();
        }

        private static string BuildFolderDiagnostics(string label, string path)
        {
            string safeLabel = label ?? string.Empty;
            string safePath = path ?? string.Empty;
            return safeLabel + "Present=" + (!string.IsNullOrWhiteSpace(safePath)).ToString()
                + ", " + safeLabel + "Length=" + safePath.Length.ToString()
                + ", " + safeLabel + "Exists=" + SafeDirectoryExists(safePath).ToString();
        }

        private static string SafeGetExtension(string path)
        {
            try
            {
                return System.IO.Path.GetExtension(path ?? string.Empty) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool SafeFileExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            try
            {
                return System.IO.File.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        private static bool SafeDirectoryExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            try
            {
                return System.IO.Directory.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        private static string FormatElapsedSeconds(TimeSpan elapsed)
        {
            return elapsed.TotalSeconds.ToString("0.000");
        }
    }
}
