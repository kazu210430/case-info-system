using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneDisplayCoordinator
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly IDictionary<string, TaskPaneHost> _hostsByWindowKey;
        private readonly KernelCaseInteractionState _kernelCaseInteractionState;
        private readonly Logger _logger;
        private readonly TaskPaneManager.TaskPaneManagerTestHooks _testHooks;
        private readonly Func<Excel.Window, string> _safeGetWindowKey;
        private readonly Func<TaskPaneHost, string> _formatHostDescriptor;
        private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
        private readonly Func<Excel.Window, string> _formatWindowDescriptor;
        private readonly Action<string> _removeHost;

        internal TaskPaneDisplayCoordinator(
            IDictionary<string, TaskPaneHost> hostsByWindowKey,
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks,
            Func<Excel.Window, string> safeGetWindowKey,
            Func<TaskPaneHost, string> formatHostDescriptor,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<Excel.Window, string> formatWindowDescriptor,
            Action<string> removeHost)
        {
            _hostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _testHooks = testHooks;
            _safeGetWindowKey = safeGetWindowKey ?? throw new ArgumentNullException(nameof(safeGetWindowKey));
            _formatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
            _formatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
            _formatWindowDescriptor = formatWindowDescriptor ?? throw new ArgumentNullException(nameof(formatWindowDescriptor));
            _removeHost = removeHost ?? throw new ArgumentNullException(nameof(removeHost));
        }

        internal bool TryShowExistingPane(ExcelInteropService excelInteropService, Excel.Workbook workbook, Excel.Window window, string reason)
        {
            string windowKey = _safeGetWindowKey(window);
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                return false;
            }

            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                return false;
            }

            string workbookFullName = workbook == null || excelInteropService == null
                ? string.Empty
                : excelInteropService.GetWorkbookFullName(workbook);
            if (!string.IsNullOrWhiteSpace(workbookFullName)
                && !string.Equals(host.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            PrepareHostsBeforeShow(host);
            if (!TryShowHost(host, "TryShowExistingPane"))
            {
                _logger.Warn("TryShowExistingPane skipped because host could not be shown. reason=" + (reason ?? string.Empty) + ", windowKey=" + windowKey);
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=show-existing-pane reason="
                + (reason ?? string.Empty)
                + ", host="
                + _formatHostDescriptor(host));
            if (workbook != null)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=show-existing-pane workbook="
                    + _formatWorkbookDescriptor(workbook));
            }
            _logger.Info("TaskPane existing host shown. reason=" + (reason ?? string.Empty) + ", windowKey=" + windowKey);
            return true;
        }

        internal TaskPaneDisplayEntryState EvaluateDisplayEntryState(ExcelInteropService excelInteropService, Excel.Workbook workbook, Excel.Window window)
        {
            return TaskPaneRenderStateEvaluator.EvaluateDisplayEntryState(
                excelInteropService,
                _hostsByWindowKey,
                workbook,
                window);
        }

        internal bool HasManagedPaneForWindow(Excel.Window window)
        {
            string windowKey = _safeGetWindowKey(window);
            return !string.IsNullOrWhiteSpace(windowKey)
                && _hostsByWindowKey.ContainsKey(windowKey);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(ExcelInteropService excelInteropService, Excel.Workbook workbook, Excel.Window window)
        {
            string windowKey = _safeGetWindowKey(window);
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=visible-case-pane-check result=NoWindowKey"
                    + ", windowKeyResolved=false"
                    + ", workbookFullNameMatched=false"
                    + ", renderCurrentCheckBypassed=true"
                    + ", visibleCasePaneEarlyComplete=false"
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook)
                    + ", inputWindow="
                    + _formatWindowDescriptor(window));
                return false;
            }

            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=visible-case-pane-check result=NoHost"
                    + ", windowKeyResolved=true"
                    + ", windowKey="
                    + windowKey
                    + ", workbookFullNameMatched=false"
                    + ", renderCurrentCheckBypassed=true"
                    + ", visibleCasePaneEarlyComplete=false"
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook));
                return false;
            }

            string workbookFullName = workbook == null || excelInteropService == null
                ? string.Empty
                : excelInteropService.GetWorkbookFullName(workbook);
            WorkbookRole hostedRole = GetHostedWorkbookRole(host);
            bool hostVisible = host.IsVisible;
            if (string.IsNullOrWhiteSpace(workbookFullName)
                || !string.Equals(host.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=visible-case-pane-check result=WorkbookMismatch"
                    + ", windowKeyResolved=true"
                    + ", windowKey="
                    + windowKey
                    + ", workbookFullNameMatched=false"
                    + ", lastRenderSignaturePresent="
                    + (!string.IsNullOrWhiteSpace(host.LastRenderSignature)).ToString()
                    + ", renderCurrentCheckBypassed=true"
                    + ", visibleCasePaneEarlyComplete=false"
                    + ", host="
                    + _formatHostDescriptor(host)
                    + ", hostRole="
                    + hostedRole.ToString()
                    + ", hostVisible="
                    + hostVisible.ToString()
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook));
                return false;
            }

            bool isVisibleCasePane = hostedRole == WorkbookRole.Case && hostVisible;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=visible-case-pane-check result="
                + (isVisibleCasePane ? "VisibleCasePaneFound" : "NotVisibleOrNotCase")
                + ", windowKeyResolved=true"
                + ", windowKey="
                + windowKey
                + ", workbookFullNameMatched=true"
                + ", lastRenderSignaturePresent="
                + (!string.IsNullOrWhiteSpace(host.LastRenderSignature)).ToString()
                + ", renderCurrentCheckBypassed=true"
                + ", visibleCasePaneEarlyComplete="
                + isVisibleCasePane.ToString()
                + ", host="
                + _formatHostDescriptor(host)
                + ", hostRole="
                + hostedRole.ToString()
                + ", hostVisible="
                + hostVisible.ToString());
            return isVisibleCasePane;
        }

        internal void HideAll()
        {
            foreach (TaskPaneHost host in new List<TaskPaneHost>(_hostsByWindowKey.Values))
            {
                SafeHideHost(host, "HideAll");
            }
        }

        internal void HideKernelPanes()
        {
            foreach (TaskPaneHost host in new List<TaskPaneHost>(_hostsByWindowKey.Values))
            {
                if (host.Control is KernelNavigationControl)
                {
                    SafeHideHost(host, "HideKernelPanes");
                }
            }
        }

        internal void HideAllExcept(string activeWindowKey)
        {
            foreach (TaskPaneHost host in new List<TaskPaneHost>(_hostsByWindowKey.Values))
            {
                if (!string.Equals(host.WindowKey, activeWindowKey, StringComparison.OrdinalIgnoreCase))
                {
                    SafeHideHost(host, "HideAllExcept");
                }
            }
        }

        internal void PrepareHostsBeforeShow(TaskPaneHost host)
        {
            if (host == null)
            {
                return;
            }

            TaskPaneHostPreparationAction action = TaskPaneHostPreparationPolicy.Decide(
                _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                host.Control is DocumentButtonsControl);

            if (action == TaskPaneHostPreparationAction.None)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=prepare-hosts decision=None"
                    + ", host="
                    + _formatHostDescriptor(host));
                return;
            }

            if (action == TaskPaneHostPreparationAction.HideNonCaseHostsExceptActiveWindow)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=prepare-hosts decision=HideNonCaseHostsExceptActiveWindow"
                    + ", host="
                    + _formatHostDescriptor(host));
                HideNonCaseHostsExcept(host.WindowKey);
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=prepare-hosts decision=HideAllExcept"
                + ", host="
                + _formatHostDescriptor(host));
            HideAllExcept(host.WindowKey);
        }

        internal void HidePaneForWindow(Excel.Window window)
        {
            string windowKey = _safeGetWindowKey(window);
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                return;
            }

            if (_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                SafeHideHost(host, "HidePaneForWindow");
            }
        }

        internal bool TryShowHost(TaskPaneHost host, string reason)
        {
            if (host == null)
            {
                return false;
            }

            if (_testHooks != null && _testHooks.TryShowHost != null)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=show-pane-test-hook reason="
                    + (reason ?? string.Empty)
                    + ", host="
                    + _formatHostDescriptor(host));
                return _testHooks.TryShowHost(host.WindowKey, reason ?? string.Empty);
            }

            try
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=show-pane reason="
                    + (reason ?? string.Empty)
                    + ", host="
                    + _formatHostDescriptor(host));
                host.Show();
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("TaskPane host show failed. reason=" + (reason ?? string.Empty) + ", windowKey=" + host.WindowKey, ex);
                _removeHost(host.WindowKey);
                return false;
            }
        }

        internal void PrepareTargetWindowForForcedRefresh(Excel.Window targetWindow)
        {
            string windowKey = _safeGetWindowKey(targetWindow);
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                return;
            }

            if (_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                InvalidateHostRenderStateForForcedRefresh(host);
            }
        }

        internal void InvalidateHostRenderStateForForcedRefresh(TaskPaneHost host)
        {
            if (host == null)
            {
                return;
            }

            host.LastRenderSignature = string.Empty;
        }

        private void HideNonCaseHostsExcept(string activeWindowKey)
        {
            foreach (TaskPaneHost host in new List<TaskPaneHost>(_hostsByWindowKey.Values))
            {
                if (string.Equals(host.WindowKey, activeWindowKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (host.Control is DocumentButtonsControl)
                {
                    continue;
                }

                SafeHideHost(host, "HideNonCaseHostsExcept");
            }
        }

        private void SafeHideHost(TaskPaneHost host, string reason)
        {
            if (host == null)
            {
                return;
            }

            try
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=hide-pane reason="
                    + (reason ?? string.Empty)
                    + ", host="
                    + _formatHostDescriptor(host));
                _testHooks?.OnHideHost?.Invoke(host.WindowKey, reason ?? string.Empty);
                host.Hide();
            }
            catch (Exception ex)
            {
                _logger.Error("TaskPane host hide failed. reason=" + (reason ?? string.Empty) + ", windowKey=" + host.WindowKey, ex);
                _removeHost(host.WindowKey);
            }
        }

        private static WorkbookRole GetHostedWorkbookRole(TaskPaneHost host)
        {
            if (host == null || host.Control == null)
            {
                return WorkbookRole.Unknown;
            }

            if (host.Control is DocumentButtonsControl)
            {
                return WorkbookRole.Case;
            }

            if (host.Control is KernelNavigationControl)
            {
                return WorkbookRole.Kernel;
            }

            if (host.Control is AccountingNavigationControl)
            {
                return WorkbookRole.Accounting;
            }

            return WorkbookRole.Unknown;
        }
    }
}
