using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // Owns TaskPane control construction and ActionInvoked binding for each role.
    // Event unbinding is not explicit in current-state; handler lifetime is still coupled to host/control disposal.
    internal sealed class TaskPaneHostFactory
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly ThisAddIn _addIn;
        private readonly Logger _logger;
        private readonly Func<TaskPaneHost, string> _formatHostDescriptor;
        private readonly Action<string, KernelNavigationActionEventArgs> _handleKernelActionInvoked;
        private readonly Action<string, AccountingNavigationActionEventArgs> _handleAccountingActionInvoked;
        private readonly Action<string, DocumentButtonsControl, TaskPaneActionEventArgs> _handleCaseActionInvoked;

        internal TaskPaneHostFactory(
            ThisAddIn addIn,
            Logger logger,
            Func<TaskPaneHost, string> formatHostDescriptor,
            Action<string, KernelNavigationActionEventArgs> handleKernelActionInvoked,
            Action<string, AccountingNavigationActionEventArgs> handleAccountingActionInvoked,
            Action<string, DocumentButtonsControl, TaskPaneActionEventArgs> handleCaseActionInvoked)
        {
            _addIn = addIn;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _formatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
            _handleKernelActionInvoked = handleKernelActionInvoked ?? throw new ArgumentNullException(nameof(handleKernelActionInvoked));
            _handleAccountingActionInvoked = handleAccountingActionInvoked ?? throw new ArgumentNullException(nameof(handleAccountingActionInvoked));
            _handleCaseActionInvoked = handleCaseActionInvoked ?? throw new ArgumentNullException(nameof(handleCaseActionInvoked));
        }

        internal TaskPaneHost CreateHost(string windowKey, Excel.Window window, WorkbookRole role, out string paneRoleName)
        {
            // Current-state boundary: binding happens here, while concrete pane lifetime is owned by TaskPaneHost and VSTO create/remove stays in ThisAddIn.
            if (role == WorkbookRole.Kernel)
            {
                var kernelControl = new KernelNavigationControl();
                kernelControl.ActionInvoked += (sender, e) => _handleKernelActionInvoked(windowKey, e);
                var host = new TaskPaneHost(_addIn, window, kernelControl, kernelControl, windowKey);
                paneRoleName = "Kernel";
                LogHostCreated(host, paneRoleName);
                return host;
            }

            if (role == WorkbookRole.Accounting)
            {
                var accountingControl = new AccountingNavigationControl();
                accountingControl.ActionInvoked += (sender, e) => _handleAccountingActionInvoked(windowKey, e);
                var host = new TaskPaneHost(_addIn, window, accountingControl, accountingControl, windowKey);
                paneRoleName = "Accounting";
                LogHostCreated(host, paneRoleName);
                return host;
            }

            var caseControl = new DocumentButtonsControl();
            var caseHost = new TaskPaneHost(_addIn, window, caseControl, caseControl, windowKey);
            caseControl.ActionInvoked += (sender, e) => _handleCaseActionInvoked(windowKey, caseControl, e);
            paneRoleName = "Case";
            LogHostCreated(caseHost, paneRoleName);
            return caseHost;
        }

        private void LogHostCreated(TaskPaneHost host, string paneRoleName)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=create-host host="
                + _formatHostDescriptor(host)
                + ", paneRole="
                + (paneRoleName ?? string.Empty));
            _logger.Info("TaskPane host created. role=" + (paneRoleName ?? string.Empty) + ", windowKey=" + (host?.WindowKey ?? string.Empty));
        }
    }
}
