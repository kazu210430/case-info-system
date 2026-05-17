using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentNamePromptService
    {
        private readonly ExcelInteropService _excelInteropService;
        private readonly ICaseCacheDocumentTemplateReader _caseCacheDocumentTemplateReader;
        private readonly Logger _logger;

        /// <summary>
        internal DocumentNamePromptService(
            ExcelInteropService excelInteropService,
            ICaseCacheDocumentTemplateReader caseCacheDocumentTemplateReader,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _caseCacheDocumentTemplateReader = caseCacheDocumentTemplateReader ?? throw new ArgumentNullException(nameof(caseCacheDocumentTemplateReader));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal bool TryPrepare(Excel.Workbook workbook, string key, out DocumentNameOverrideScope scope)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            scope = null;

            string initialDocumentName = FindDocumentCaptionByKey(workbook, key);
            string finalDocumentName;
            bool accepted;
            using (ExcelWindowOwner owner = ExcelWindowOwner.From(ResolvePreferredWindow(workbook)))
            {
                accepted = DocumentNamePromptForm.TryPrompt(owner, initialDocumentName, out finalDocumentName);
            }

            if (!accepted)
            {
                _logger.Info(
                    "Document name prompt cancelled. key="
                    + (key ?? string.Empty)
                    + ", promptResult=Cancelled, "
                    + BuildDocumentNameDiagnostics("initialDocumentName", initialDocumentName));
                return false;
            }

            scope = new DocumentNameOverrideScope(_excelInteropService, workbook, _logger, finalDocumentName);
            _logger.Info(
                "Document name prompt accepted. key="
                + (key ?? string.Empty)
                + ", promptResult=Accepted, "
                + BuildDocumentNameDiagnostics("initialDocumentName", initialDocumentName)
                + ", "
                + BuildDocumentNameDiagnostics("finalDocumentName", finalDocumentName));
            return true;
        }

        /// <summary>
        private string FindDocumentCaptionByKey(Excel.Workbook workbook, string key)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            return _caseCacheDocumentTemplateReader.TryEnsurePromotedCaseCacheThenResolve(workbook, key, out Domain.DocumentTemplateLookupResult lookupResult)
                ? lookupResult.DocumentName
                : string.Empty;
        }

        /// <summary>
        private Excel.Window ResolvePreferredWindow(Excel.Workbook workbook)
        {
            Excel.Window activeWindow = _excelInteropService.GetActiveWindow();
            Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();

            if (activeWindow != null
                && activeWorkbook != null
                && string.Equals(
                    _excelInteropService.GetWorkbookFullName(activeWorkbook),
                    _excelInteropService.GetWorkbookFullName(workbook),
                    StringComparison.OrdinalIgnoreCase))
            {
                return activeWindow;
            }

            return _excelInteropService.GetFirstVisibleWindow(workbook);
        }

        private static string BuildDocumentNameDiagnostics(string label, string documentName)
        {
            string safeLabel = label ?? string.Empty;
            string safeDocumentName = documentName ?? string.Empty;
            return safeLabel + "Provided=" + (!string.IsNullOrWhiteSpace(safeDocumentName)).ToString()
                + ", " + safeLabel + "Length=" + safeDocumentName.Length.ToString()
                + ", " + safeLabel + "TrimmedLength=" + safeDocumentName.Trim().Length.ToString();
        }
    }
}
