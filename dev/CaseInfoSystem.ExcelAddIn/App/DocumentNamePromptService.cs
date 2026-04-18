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
        private readonly TaskPaneSnapshotCacheService _taskPaneSnapshotCacheService;
        private readonly Logger _logger;

        /// <summary>
        internal DocumentNamePromptService(
            ExcelInteropService excelInteropService,
            TaskPaneSnapshotCacheService taskPaneSnapshotCacheService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _taskPaneSnapshotCacheService = taskPaneSnapshotCacheService ?? throw new ArgumentNullException(nameof(taskPaneSnapshotCacheService));
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
                    "Document name prompt cancelled. workbook="
                    + _excelInteropService.GetWorkbookName(workbook)
                    + ", key="
                    + (key ?? string.Empty));
                return false;
            }

            scope = new DocumentNameOverrideScope(_excelInteropService, workbook, _logger, finalDocumentName);
            _logger.Info(
                "Document name prompt accepted. workbook="
                + _excelInteropService.GetWorkbookName(workbook)
                + ", key="
                + (key ?? string.Empty)
                + ", finalDocumentName="
                + finalDocumentName);
            return true;
        }

        /// <summary>
        private string FindDocumentCaptionByKey(Excel.Workbook workbook, string key)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            return _taskPaneSnapshotCacheService.TryGetDocInfoFromCache(workbook, key, out string _, out string documentName)
                ? documentName
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
    }
}
