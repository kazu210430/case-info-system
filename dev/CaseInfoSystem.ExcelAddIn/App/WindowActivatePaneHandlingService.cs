using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WindowActivatePaneHandlingService
    {
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly TaskPaneManager _taskPaneManager;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly IExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        internal WindowActivatePaneHandlingService(
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            TaskPaneManager taskPaneManager,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane,
            IExcelInteropService excelInteropService,
            Logger logger)
        {
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _taskPaneManager = taskPaneManager;
            _refreshTaskPane = refreshTaskPane;
            _excelInteropService = excelInteropService;
            _logger = logger;
        }

        internal void Handle(Excel.Workbook workbook, Excel.Window window)
        {
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate();
            string reason = request.ToReasonString();
            if (Globals.ThisAddIn != null && Globals.ThisAddIn.ShouldIgnoreWindowActivateDuringCaseProtection(workbook, window))
            {
                return;
            }

            _handleExternalWorkbookDetected?.Invoke(workbook, reason);
            if (_shouldSuppressCasePaneRefresh != null && _shouldSuppressCasePaneRefresh(reason, workbook))
            {
                return;
            }

            _refreshTaskPane?.Invoke(request, workbook, window);
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            return window == null ? "0" : window.Hwnd.ToString();
        }
    }
}
