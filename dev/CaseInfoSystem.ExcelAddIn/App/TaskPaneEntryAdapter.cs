using System;
using System.Globalization;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneEntryAdapter
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Logger _logger;
        private readonly Func<string> _formatActiveExcelState;
        private readonly Func<Excel.Window, UserControl, CustomTaskPane> _createTaskPane;
        private readonly Action<CustomTaskPane> _removeTaskPane;
        private ExcelInteropService _excelInteropService;
        private WorkbookRoleResolver _workbookRoleResolver;
        private TaskPaneManager _taskPaneManager;
        private TaskPaneRefreshOrchestrationService _taskPaneRefreshOrchestrationService;
        private int _kernelFlickerTraceRefreshCallSequence;

        internal TaskPaneEntryAdapter(
            Logger logger,
            Func<string> formatActiveExcelState,
            Func<Excel.Window, UserControl, CustomTaskPane> createTaskPane,
            Action<CustomTaskPane> removeTaskPane)
        {
            _logger = logger;
            _formatActiveExcelState = formatActiveExcelState;
            _createTaskPane = createTaskPane ?? throw new ArgumentNullException(nameof(createTaskPane));
            _removeTaskPane = removeTaskPane ?? throw new ArgumentNullException(nameof(removeTaskPane));
        }

        internal void Configure(
            ExcelInteropService excelInteropService,
            WorkbookRoleResolver workbookRoleResolver,
            TaskPaneManager taskPaneManager,
            TaskPaneRefreshOrchestrationService taskPaneRefreshOrchestrationService)
        {
            _excelInteropService = excelInteropService;
            _workbookRoleResolver = workbookRoleResolver;
            _taskPaneManager = taskPaneManager;
            _taskPaneRefreshOrchestrationService = taskPaneRefreshOrchestrationService;
        }

        internal void RequestTaskPaneDisplayForTargetWindow(TaskPaneDisplayRequest request, Excel.Workbook workbook, Excel.Window targetWindow)
        {
            if (request != null && request.RefreshIntent == TaskPaneDisplayRefreshIntent.ForceRefresh)
            {
                _taskPaneManager?.PrepareTargetWindowForForcedRefresh(targetWindow);
            }

            TaskPaneDisplayEntryDecision displayEntryDecision = PaneDisplayPolicy.Decide(
                request,
                _taskPaneManager,
                _workbookRoleResolver,
                workbook,
                targetWindow);
            LogTaskPaneDisplayEntryDecision(request, displayEntryDecision, workbook, targetWindow);
            switch (displayEntryDecision.Result)
            {
                case PaneDisplayPolicyResult.ShowExisting:
                    _taskPaneManager?.TryShowExistingPane(workbook, targetWindow, "DisplayRequest.ShowExisting");
                    return;

                case PaneDisplayPolicyResult.Hide:
                    _taskPaneManager?.HidePaneForWindow(targetWindow);
                    return;

                case PaneDisplayPolicyResult.Reject:
                    return;
            }

            RefreshTaskPane(request, workbook, targetWindow);
        }

        internal void RefreshTaskPane(TaskPaneDisplayRequest request, Excel.Workbook workbook, Excel.Window window)
        {
            string reason = request == null ? string.Empty : request.ToReasonString();
            RefreshTaskPane(reason, workbook, window, request);
        }

        internal void RefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            RefreshTaskPane(reason, workbook, window, request: null);
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneRefreshOrchestrationService == null
                ? TaskPaneRefreshAttemptResult.Skipped("taskPaneRefreshOrchestrationServiceUnavailable")
                : _taskPaneRefreshOrchestrationService.TryRefreshTaskPane(reason, workbook, window);
        }

        internal bool IsTaskPaneRefreshSucceeded(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return TryRefreshTaskPane(reason, workbook, window).IsRefreshSucceeded;
        }

        internal void RefreshActiveTaskPane(string reason)
        {
            _taskPaneRefreshOrchestrationService?.RefreshActiveTaskPane(reason);
        }

        internal void ScheduleActiveTaskPaneRefresh(string reason)
        {
            _taskPaneRefreshOrchestrationService?.ScheduleActiveTaskPaneRefresh(reason);
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
            _taskPaneRefreshOrchestrationService?.ScheduleWorkbookTaskPaneRefresh(workbook, reason);
        }

        internal void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            _taskPaneRefreshOrchestrationService?.ShowWorkbookTaskPaneWhenReady(workbook, reason);
        }

        internal Excel.Window ResolveWorkbookPaneWindow(Excel.Workbook workbook, string reason, bool activateWorkbook)
        {
            return _taskPaneRefreshOrchestrationService == null
                ? null
                : _taskPaneRefreshOrchestrationService.ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneManager != null
                && _taskPaneManager.HasVisibleCasePaneForWorkbookWindow(workbook, window);
        }

        internal CustomTaskPane CreateTaskPane(Excel.Window window, UserControl control)
        {
            CustomTaskPane pane = _createTaskPane(window, control);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            return pane;
        }

        internal void RemoveTaskPane(CustomTaskPane pane)
        {
            if (pane == null)
            {
                return;
            }

            _removeTaskPane(pane);
        }

        internal void StopPendingPaneRefreshTimer()
        {
            _taskPaneRefreshOrchestrationService?.StopPendingPaneRefreshTimer();
        }

        internal void DisposeAll()
        {
            _taskPaneManager?.DisposeAll();
        }

        private void RefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window, TaskPaneDisplayRequest request)
        {
            int refreshCallId = ++_kernelFlickerTraceRefreshCallSequence;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=refresh-call-start refreshCallId="
                + refreshCallId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", activeState="
                + FormatActiveExcelState()
                + FormatDisplayRequestTraceFields(request));
            TaskPaneRefreshAttemptResult result = request == null
                ? TryRefreshTaskPane(reason, workbook, window)
                : _taskPaneRefreshOrchestrationService.TryRefreshTaskPane(request, workbook, window);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=refresh-call-end refreshCallId="
                + refreshCallId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", result="
                + (result == null ? "null" : result.IsRefreshSucceeded.ToString())
                + FormatDisplayRequestTraceFields(request));
        }

        private void LogTaskPaneDisplayEntryDecision(
            TaskPaneDisplayRequest request,
            TaskPaneDisplayEntryDecision decision,
            Excel.Workbook workbook,
            Excel.Window targetWindow)
        {
            if (request == null || !request.IsWindowActivateTrigger)
            {
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=window-activate-display-entry-decision reason="
                + request.ToReasonString()
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", displayEntryResult="
                + (decision == null ? PaneDisplayPolicyResult.Reject.ToString() : decision.Result.ToString())
                + ", displayCompletionOutcome=False"
                + ", recoveryOwner=False"
                + ", foregroundGuaranteeOwner=False"
                + ", hiddenExcelOwner=False"
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", targetWindow="
                + FormatWindowDescriptor(targetWindow)
                + FormatDisplayEntryStateTraceFields(decision == null ? null : decision.State)
                + FormatDisplayRequestTraceFields(request));
        }

        private string FormatActiveExcelState()
        {
            return _formatActiveExcelState == null ? string.Empty : _formatActiveExcelState();
        }

        private static string FormatDisplayEntryStateTraceFields(TaskPaneDisplayEntryState state)
        {
            if (state == null)
            {
                return ", displayEntryState=null";
            }

            return ", displayEntryState=present"
                + ", hasTargetWindow=" + state.HasTargetWindow.ToString()
                + ", hasResolvableWindowKey=" + state.HasResolvableWindowKey.ToString()
                + ", hasManagedPane=" + state.HasManagedPane.ToString()
                + ", hasExistingHost=" + state.HasExistingHost.ToString()
                + ", isSameWorkbook=" + state.IsSameWorkbook.ToString()
                + ", isRenderSignatureCurrent=" + state.IsRenderSignatureCurrent.ToString();
        }

        private static string FormatDisplayRequestTraceFields(TaskPaneDisplayRequest request)
        {
            if (request == null)
            {
                return string.Empty;
            }

            string details =
                ", displayRequestSource=" + request.Source.ToString()
                + ", displayRequestRefreshIntent=" + request.RefreshIntent.ToString()
                + ", displayTriggerReason=" + request.ToReasonString();
            if (!request.IsWindowActivateTrigger)
            {
                return details;
            }

            WindowActivateTaskPaneTriggerFacts facts = request.WindowActivateTriggerFacts;
            return details
                + ", windowActivateTriggerRole=TaskPaneDisplayRefreshTrigger"
                + ", windowActivateRecoveryOwner=False"
                + ", windowActivateForegroundGuaranteeOwner=False"
                + ", windowActivateHiddenExcelOwner=False"
                + ", windowActivateCaptureOwner=" + (facts == null ? string.Empty : facts.CaptureOwner)
                + ", windowActivateWorkbookPresent=" + (facts != null && facts.HasWorkbook).ToString()
                + ", windowActivateWindowPresent=" + (facts != null && facts.HasWindow).ToString()
                + ", windowActivateWindowHwnd=" + (facts == null ? string.Empty : facts.WindowHwnd);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        private string SafeWorkbookName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookName(workbook);
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\"";
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowCaption(Excel.Window window)
        {
            try
            {
                if (window == null)
                {
                    return string.Empty;
                }

                dynamic lateBoundWindow = window;
                return Convert.ToString(lateBoundWindow.Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
