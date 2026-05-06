using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelTemplateFolderPathResolver
    {
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string WordTemplateDirectoryPropertyName = "WORD_TEMPLATE_DIR";
        private const string TemplateFolderName = "雛形";

        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Func<Excel.Workbook, string, string> _tryGetDocumentProperty;
        private readonly Func<Excel.Workbook, string> _getWorkbookPath;

        internal KernelTemplateFolderPathResolver(
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService)
            : this(
                pathCompatibilityService,
                (workbook, propertyName) => excelInteropService == null ? string.Empty : excelInteropService.TryGetDocumentProperty(workbook, propertyName),
                workbook => excelInteropService == null ? string.Empty : excelInteropService.GetWorkbookPath(workbook))
        {
            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }
        }

        internal KernelTemplateFolderPathResolver(
            PathCompatibilityService pathCompatibilityService,
            Func<Excel.Workbook, string, string> tryGetDocumentProperty,
            Func<Excel.Workbook, string> getWorkbookPath)
        {
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _tryGetDocumentProperty = tryGetDocumentProperty ?? throw new ArgumentNullException(nameof(tryGetDocumentProperty));
            _getWorkbookPath = getWorkbookPath ?? throw new ArgumentNullException(nameof(getWorkbookPath));
        }

        internal string ResolveDirectoryForOpen(Excel.Workbook kernelWorkbook)
        {
            if (kernelWorkbook == null)
            {
                throw new ArgumentNullException(nameof(kernelWorkbook));
            }

            string configuredDirectory = ResolveConfiguredDirectory(kernelWorkbook);
            if (!string.IsNullOrWhiteSpace(configuredDirectory)
                && _pathCompatibilityService.DirectoryExistsSafe(configuredDirectory))
            {
                return configuredDirectory;
            }

            return ResolveDefaultDirectory(kernelWorkbook);
        }

        internal string ResolveConfiguredDirectory(Excel.Workbook kernelWorkbook)
        {
            if (kernelWorkbook == null)
            {
                throw new ArgumentNullException(nameof(kernelWorkbook));
            }

            return _pathCompatibilityService.NormalizePath(
                _tryGetDocumentProperty(kernelWorkbook, WordTemplateDirectoryPropertyName));
        }

        internal string ResolveDefaultDirectory(Excel.Workbook kernelWorkbook)
        {
            if (kernelWorkbook == null)
            {
                throw new ArgumentNullException(nameof(kernelWorkbook));
            }

            string systemRoot = ResolveSystemRoot(kernelWorkbook);
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                return string.Empty;
            }

            return _pathCompatibilityService.NormalizePath(
                _pathCompatibilityService.CombinePath(systemRoot, TemplateFolderName));
        }

        private string ResolveSystemRoot(Excel.Workbook kernelWorkbook)
        {
            string systemRoot = _pathCompatibilityService.NormalizePath(
                _tryGetDocumentProperty(kernelWorkbook, SystemRootPropertyName));
            if (!string.IsNullOrWhiteSpace(systemRoot))
            {
                return systemRoot;
            }

            return _pathCompatibilityService.NormalizePath(_getWorkbookPath(kernelWorkbook));
        }
    }
}
