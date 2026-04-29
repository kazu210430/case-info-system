using System;
using System.Diagnostics;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentTemplateResolver
    {
        private const string WordTemplateDirectoryPropertyName = "WORD_TEMPLATE_DIR";
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string DefaultTemplateFolderName = "\u96DB\u5F62";

        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly DocumentTemplateLookupService _documentTemplateLookupService;
        private readonly Logger _logger;

        /// <summary>
        internal DocumentTemplateResolver(
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            DocumentTemplateLookupService documentTemplateLookupService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _documentTemplateLookupService = documentTemplateLookupService ?? throw new ArgumentNullException(nameof(documentTemplateLookupService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal DocumentTemplateSpec Resolve(Excel.Workbook workbook, string key)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            string normalizedKey = (key ?? string.Empty).Trim();
            if (normalizedKey.Length == 0)
            {
                return null;
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            if (!_documentTemplateLookupService.TryResolveWithMasterFallback(workbook, normalizedKey, out DocumentTemplateLookupResult lookupResult))
            {
                _logger.Info("DocumentTemplateResolver could not resolve key. key=" + normalizedKey);
                return null;
            }

            string templateFileName = lookupResult.TemplateFileName ?? string.Empty;
            string documentName = lookupResult.DocumentName ?? string.Empty;
            string templateDirectory = ResolveTemplateDirectory(workbook);
            string templatePath = templateFileName.Length == 0 || templateDirectory.Length == 0
                ? string.Empty
                : _pathCompatibilityService.NormalizePath(_pathCompatibilityService.CombinePath(templateDirectory, templateFileName));

            _logger.Debug(
                "DocumentTemplateResolver.Resolve",
                "Completed elapsed=" + FormatElapsedSeconds(stopwatch.Elapsed)
                + " key=" + normalizedKey
                + " source=" + lookupResult.ResolutionSource.ToString()
                + " templateFile=" + templateFileName
                + " templatePath=" + templatePath);

            return new DocumentTemplateSpec
            {
                Key = normalizedKey,
                DocumentName = documentName,
                TemplateFileName = templateFileName,
                TemplatePath = templatePath,
                ActionKind = "doc",
                ResolutionSource = lookupResult.ResolutionSource
            };
        }

        /// <summary>
        internal static bool IsSupportedWordTemplate(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null || string.IsNullOrWhiteSpace(templateSpec.TemplatePath))
            {
                return false;
            }

            string extension = Path.GetExtension(templateSpec.TemplatePath) ?? string.Empty;
            return string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".dotx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".dotm", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        internal bool TemplateExists(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null || string.IsNullOrWhiteSpace(templateSpec.TemplatePath))
            {
                return false;
            }

            try
            {
                return _pathCompatibilityService.FileExistsSafe(templateSpec.TemplatePath);
            }
            catch (Exception ex)
            {
                _logger.Error("DocumentTemplateResolver.TemplateExists failed.", ex);
                return false;
            }
        }

        /// <summary>
        private string ResolveTemplateDirectory(Excel.Workbook workbook)
        {
            string configuredDirectory = (_excelInteropService.TryGetDocumentProperty(workbook, WordTemplateDirectoryPropertyName) ?? string.Empty).Trim();
            if (configuredDirectory.Length > 0)
            {
                return _pathCompatibilityService.NormalizePath(configuredDirectory);
            }

            string systemRoot = (_excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName) ?? string.Empty).Trim();
            if (systemRoot.Length == 0)
            {
                return string.Empty;
            }

            return _pathCompatibilityService.NormalizePath(_pathCompatibilityService.CombinePath(systemRoot, DefaultTemplateFolderName));
        }

        private static string FormatElapsedSeconds(TimeSpan elapsed)
        {
            return elapsed.TotalSeconds.ToString("0.000");
        }
    }
}
