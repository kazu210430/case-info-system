using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WindowActivatePaneHandlingService
    {
        private const string EventName = "WindowActivate";
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly TaskPaneManager _taskPaneManager;
        private readonly Action<string, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly IExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        internal WindowActivatePaneHandlingService(
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            TaskPaneManager taskPaneManager,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPane,
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
            _handleExternalWorkbookDetected?.Invoke(workbook, EventName);
            if (_shouldSuppressCasePaneRefresh != null && _shouldSuppressCasePaneRefresh(EventName, workbook))
            {
                return;
            }

            if (_taskPaneManager != null && _taskPaneManager.TryShowExistingPane(workbook, window, EventName))
            {
                _logger?.Info(
                    "WindowActivate reused existing pane. workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook))
                    + ", windowHwnd="
                    + SafeWindowHwnd(window));
                return;
            }

            _refreshTaskPane?.Invoke(EventName, workbook, window);
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            return window == null ? "0" : window.Hwnd.ToString();
        }
    }
}
