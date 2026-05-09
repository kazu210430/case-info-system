using System;
using System.Diagnostics;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WindowActivateDownstreamObservation
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
        private readonly Func<Excel.Window, string> _formatWindowDescriptor;
        private readonly Func<string> _formatActiveState;

        internal WindowActivateDownstreamObservation(
            Logger logger,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<Excel.Window, string> formatWindowDescriptor,
            Func<string> formatActiveState)
        {
            _logger = logger;
            _formatWorkbookDescriptor = formatWorkbookDescriptor ?? (_ => string.Empty);
            _formatWindowDescriptor = formatWindowDescriptor ?? (_ => string.Empty);
            _formatActiveState = formatActiveState ?? (() => string.Empty);
        }

        internal void LogStart(
            TaskPaneDisplayRequest displayRequest,
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            int refreshAttemptId)
        {
            if (!IsWindowActivateDisplayRequest(displayRequest))
            {
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-start refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", windowActivateDispatchStatus=Dispatched"
                + ", activationAttempt=NotAttempted"
                + ", downstreamRecoveryDelegated=False"
                + ", displayCompletionOutcome=False"
                + ", recoveryOwner=False"
                + ", foregroundGuaranteeOwner=False"
                + ", hiddenExcelOwner=False"
                + ", workbook="
                + _formatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + _formatWindowDescriptor(window)
                + ", activeState="
                + _formatActiveState()
                + FormatDisplayRequestTraceFields(displayRequest));
        }

        internal void LogOutcome(
            TaskPaneDisplayRequest displayRequest,
            string reason,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch,
            int refreshAttemptId,
            string completionSource)
        {
            if (!IsWindowActivateDisplayRequest(displayRequest))
            {
                return;
            }

            VisibilityRecoveryOutcome visibilityOutcome = attemptResult == null ? null : attemptResult.VisibilityRecoveryOutcome;
            RefreshSourceSelectionOutcome refreshSourceOutcome = attemptResult == null ? null : attemptResult.RefreshSourceSelectionOutcome;
            RebuildFallbackOutcome rebuildOutcome = attemptResult == null ? null : attemptResult.RebuildFallbackOutcome;
            ForegroundGuaranteeOutcome foregroundOutcome = attemptResult == null ? null : attemptResult.ForegroundGuaranteeOutcome;
            bool downstreamRecoveryDelegated = (attemptResult != null && attemptResult.PreContextRecoveryAttempted)
                || (foregroundOutcome != null && foregroundOutcome.WasExecutionAttempted);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-outcome refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", completionSource="
                + (completionSource ?? string.Empty)
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", windowActivateDispatchStatus=Dispatched"
                + ", activationAttempt="
                + (downstreamRecoveryDelegated
                    ? WindowActivateActivationAttempt.Delegated.ToString()
                    : WindowActivateActivationAttempt.NotAttempted.ToString())
                + ", downstreamRecoveryDelegated="
                + downstreamRecoveryDelegated.ToString()
                + ", displayCompletionOutcome=False"
                + ", recoveryOwner=False"
                + ", foregroundGuaranteeOwner=False"
                + ", hiddenExcelOwner=False"
                + ", refreshSucceeded="
                + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ", refreshSkipped="
                + (attemptResult != null && attemptResult.WasSkipped).ToString()
                + ", contextRejected="
                + (attemptResult != null && attemptResult.WasContextRejected).ToString()
                + ", paneVisible="
                + (attemptResult != null && attemptResult.IsPaneVisible).ToString()
                + ", refreshCompleted="
                + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ", visibilityRecoveryStatus="
                + (visibilityOutcome == null ? VisibilityRecoveryOutcomeStatus.Unknown.ToString() : visibilityOutcome.Status.ToString())
                + ", visibilityRecoveryReason="
                + (visibilityOutcome == null ? string.Empty : visibilityOutcome.Reason)
                + ", refreshSourceStatus="
                + (refreshSourceOutcome == null ? RefreshSourceSelectionOutcomeStatus.Unknown.ToString() : refreshSourceOutcome.Status.ToString())
                + ", refreshSourceSelectedSource="
                + (refreshSourceOutcome == null ? TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None.ToString() : refreshSourceOutcome.SelectedSource.ToString())
                + ", rebuildFallbackStatus="
                + (rebuildOutcome == null ? RebuildFallbackOutcomeStatus.Unknown.ToString() : rebuildOutcome.Status.ToString())
                + ", foregroundGuaranteeStatus="
                + (foregroundOutcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : foregroundOutcome.Status.ToString())
                + ", foregroundExecutionAttempted="
                + (foregroundOutcome != null && foregroundOutcome.WasExecutionAttempted).ToString()
                + ", preContextFullRecoveryAttempted="
                + (attemptResult != null && attemptResult.PreContextRecoveryAttempted).ToString()
                + ", elapsedMs="
                + (stopwatch == null ? 0 : stopwatch.ElapsedMilliseconds).ToString(CultureInfo.InvariantCulture)
                + FormatDisplayRequestTraceFields(displayRequest));
        }

        internal static bool IsWindowActivateDisplayRequest(TaskPaneDisplayRequest displayRequest)
        {
            return displayRequest != null && displayRequest.IsWindowActivateTrigger;
        }

        internal static string FormatDisplayRequestTraceFields(TaskPaneDisplayRequest displayRequest)
        {
            if (displayRequest == null)
            {
                return string.Empty;
            }

            string details =
                ", displayRequestSource=" + displayRequest.Source.ToString()
                + ", displayRequestRefreshIntent=" + displayRequest.RefreshIntent.ToString()
                + ", displayTriggerReason=" + displayRequest.ToReasonString();
            if (!displayRequest.IsWindowActivateTrigger)
            {
                return details;
            }

            WindowActivateTaskPaneTriggerFacts facts = displayRequest.WindowActivateTriggerFacts;
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
    }
}
