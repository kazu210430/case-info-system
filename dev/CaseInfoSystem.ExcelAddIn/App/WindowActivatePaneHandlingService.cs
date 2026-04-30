using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
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
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private readonly IWindowActivatePanePredicateBridge _windowActivatePanePredicateBridge;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Logger _logger;

        internal WindowActivatePaneHandlingService(
            IWindowActivatePanePredicateBridge windowActivatePanePredicateBridge,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane,
            Logger logger = null)
        {
            _windowActivatePanePredicateBridge = windowActivatePanePredicateBridge ?? throw new ArgumentNullException(nameof(windowActivatePanePredicateBridge));
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPane = refreshTaskPane;
            _logger = logger;
        }

        internal void Handle(Excel.Workbook workbook, Excel.Window window)
        {
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate();
            string reason = request.ToReasonString();
            if (_windowActivatePanePredicateBridge.ShouldIgnoreDuringCaseProtection(workbook, window))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=WindowActivatePaneHandlingService action=protection-return reason="
                    + reason
                    + ", workbookNull="
                    + (workbook == null).ToString()
                    + ", windowHwnd="
                    + SafeWindowHwnd(window));
                return;
            }

            _handleExternalWorkbookDetected?.Invoke(workbook, reason);
            if (_shouldSuppressCasePaneRefresh != null && _shouldSuppressCasePaneRefresh(reason, workbook))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=WindowActivatePaneHandlingService action=suppression-return reason="
                    + reason
                    + ", workbookNull="
                    + (workbook == null).ToString()
                    + ", windowHwnd="
                    + SafeWindowHwnd(window));
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WindowActivatePaneHandlingService action=refresh-proceed reason="
                + reason
                + ", workbookNull="
                + (workbook == null).ToString()
                + ", windowHwnd="
                + SafeWindowHwnd(window));
            _refreshTaskPane?.Invoke(request, workbook, window);
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            return window == null ? "0" : window.Hwnd.ToString();
        }
    }
}
