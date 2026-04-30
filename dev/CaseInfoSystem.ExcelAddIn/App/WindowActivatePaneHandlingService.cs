using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal interface IWindowActivatePanePredicateBridge
    {
        bool ShouldIgnoreDuringCaseProtection(Excel.Workbook workbook, Excel.Window window);
    }

    internal sealed class ThisAddInWindowActivatePanePredicateBridge : IWindowActivatePanePredicateBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInWindowActivatePanePredicateBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public bool ShouldIgnoreDuringCaseProtection(Excel.Workbook workbook, Excel.Window window)
        {
            return _addIn.ShouldIgnoreWindowActivateDuringCaseProtection(workbook, window);
        }
    }

    internal sealed class WindowActivatePaneHandlingService
    {
        private readonly IWindowActivatePanePredicateBridge _windowActivatePanePredicateBridge;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;

        internal WindowActivatePaneHandlingService(
            IWindowActivatePanePredicateBridge windowActivatePanePredicateBridge,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane)
        {
            _windowActivatePanePredicateBridge = windowActivatePanePredicateBridge ?? throw new ArgumentNullException(nameof(windowActivatePanePredicateBridge));
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPane = refreshTaskPane;
        }

        internal void Handle(Excel.Workbook workbook, Excel.Window window)
        {
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate();
            string reason = request.ToReasonString();
            if (_windowActivatePanePredicateBridge.ShouldIgnoreDuringCaseProtection(workbook, window))
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
