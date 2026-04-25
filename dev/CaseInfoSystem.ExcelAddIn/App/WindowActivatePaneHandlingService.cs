using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WindowActivatePaneHandlingService
    {
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;

        internal WindowActivatePaneHandlingService(
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane)
        {
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPane = refreshTaskPane;
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
