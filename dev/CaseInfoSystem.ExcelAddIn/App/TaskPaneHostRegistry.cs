using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneHostRegistry
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Dictionary<string, TaskPaneHost> _hostsByWindowKey;
        private readonly ThisAddIn _addIn;
        private readonly Logger _logger;
        private readonly Func<TaskPaneHost, string> _formatHostDescriptor;
        private readonly Action<string, KernelNavigationActionEventArgs> _handleKernelActionInvoked;
        private readonly Action<string, AccountingNavigationActionEventArgs> _handleAccountingActionInvoked;
        private readonly Action<string, DocumentButtonsControl, TaskPaneActionEventArgs> _handleCaseActionInvoked;

        internal TaskPaneHostRegistry(
            Dictionary<string, TaskPaneHost> hostsByWindowKey,
            ThisAddIn addIn,
            Logger logger,
            Func<TaskPaneHost, string> formatHostDescriptor,
            Action<string, KernelNavigationActionEventArgs> handleKernelActionInvoked,
            Action<string, AccountingNavigationActionEventArgs> handleAccountingActionInvoked,
            Action<string, DocumentButtonsControl, TaskPaneActionEventArgs> handleCaseActionInvoked)
        {
            _hostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
            _addIn = addIn;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _formatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
            _handleKernelActionInvoked = handleKernelActionInvoked ?? throw new ArgumentNullException(nameof(handleKernelActionInvoked));
            _handleAccountingActionInvoked = handleAccountingActionInvoked ?? throw new ArgumentNullException(nameof(handleAccountingActionInvoked));
            _handleCaseActionInvoked = handleCaseActionInvoked ?? throw new ArgumentNullException(nameof(handleCaseActionInvoked));
        }

        internal void RegisterHost(TaskPaneHost host)
        {
            if (host == null)
            {
                throw new ArgumentNullException(nameof(host));
            }

            if (_hostsByWindowKey.TryGetValue(host.WindowKey, out TaskPaneHost existingHost)
                && !ReferenceEquals(existingHost, host))
            {
                existingHost.Dispose();
            }

            _hostsByWindowKey[host.WindowKey] = host;
        }

        internal TaskPaneHost GetOrReplaceHost(string windowKey, Excel.Window window, WorkbookRole role)
        {
            if (_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost existingHost))
            {
                bool roleMatches =
                    (role == WorkbookRole.Kernel && existingHost.Control is KernelNavigationControl)
                    || (role == WorkbookRole.Case && existingHost.Control is DocumentButtonsControl)
                    || (role == WorkbookRole.Accounting && existingHost.Control is AccountingNavigationControl);
                if (roleMatches)
                {
                    return existingHost;
                }

                existingHost.Dispose();
                _hostsByWindowKey.Remove(windowKey);
            }

            if (role == WorkbookRole.Kernel)
            {
                var kernelControl = new KernelNavigationControl();
                kernelControl.ActionInvoked += (sender, e) => _handleKernelActionInvoked(windowKey, e);
                var host = new TaskPaneHost(_addIn, window, kernelControl, kernelControl, windowKey);
                _hostsByWindowKey.Add(windowKey, host);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=create-host host="
                    + _formatHostDescriptor(host)
                    + ", paneRole=Kernel");
                _logger.Info("TaskPane host created. role=Kernel, windowKey=" + windowKey);
                return host;
            }

            if (role == WorkbookRole.Accounting)
            {
                var accountingControl = new AccountingNavigationControl();
                accountingControl.ActionInvoked += (sender, e) => _handleAccountingActionInvoked(windowKey, e);
                var host = new TaskPaneHost(_addIn, window, accountingControl, accountingControl, windowKey);
                _hostsByWindowKey.Add(windowKey, host);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=create-host host="
                    + _formatHostDescriptor(host)
                    + ", paneRole=Accounting");
                _logger.Info("TaskPane host created. role=Accounting, windowKey=" + windowKey);
                return host;
            }

            var caseControl = new DocumentButtonsControl();
            var caseHost = new TaskPaneHost(_addIn, window, caseControl, caseControl, windowKey);
            caseControl.ActionInvoked += (sender, e) => _handleCaseActionInvoked(windowKey, caseControl, e);
            _hostsByWindowKey.Add(windowKey, caseHost);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=create-host host="
                + _formatHostDescriptor(caseHost)
                + ", paneRole=Case");
            _logger.Info("TaskPane host created. role=Case, windowKey=" + windowKey);
            return caseHost;
        }

        internal void RemoveWorkbookPanes(string workbookFullName)
        {
            var targetKeys = new List<string>();
            foreach (KeyValuePair<string, TaskPaneHost> pair in _hostsByWindowKey)
            {
                if (string.Equals(pair.Value.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                {
                    targetKeys.Add(pair.Key);
                }
            }

            foreach (string windowKey in targetKeys)
            {
                RemoveHost(windowKey);
            }
        }

        internal void DisposeAll()
        {
            foreach (TaskPaneHost host in new List<TaskPaneHost>(_hostsByWindowKey.Values))
            {
                host.Dispose();
            }

            _hostsByWindowKey.Clear();
        }

        internal void RemoveHost(string windowKey)
        {
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                return;
            }

            if (_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=remove-host host="
                    + _formatHostDescriptor(host));
                _hostsByWindowKey.Remove(windowKey);
                try
                {
                    host.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.Error("TaskPane host dispose failed. windowKey=" + windowKey, ex);
                }
            }
        }
    }
}
