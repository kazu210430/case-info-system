using System.Globalization;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshEmitPayloadBuilder
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private const string OrchestrationSource = "TaskPaneRefreshOrchestrationService";
        private const string CompleteCreatedCaseDisplaySessionSource =
            "TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession";

        internal CaseDisplayCompletedPayload BuildCaseDisplayCompleted(
            CaseDisplayCompletedPayloadInput input)
        {
            string details =
                "reason=" + (input.Reason ?? string.Empty)
                + ",sessionId=" + input.SessionId
                + ",completionSource=" + (input.CompletionSource ?? string.Empty)
                + ",completion=" + input.AttemptResult.CompletionBasis
                + ",paneVisible=" + input.AttemptResult.IsPaneVisible.ToString()
                + ",visibilityRecoveryStatus=" + input.AttemptResult.VisibilityRecoveryOutcome.Status.ToString()
                + ",visibilityRecoveryDisplayCompletable=" + input.AttemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable.ToString()
                + ",visibilityRecoveryPaneVisible=" + input.AttemptResult.VisibilityRecoveryOutcome.IsPaneVisible.ToString()
                + ",visibilityRecoveryTargetKind=" + input.AttemptResult.VisibilityRecoveryOutcome.TargetKind.ToString()
                + ",visibilityPaneVisibleSource=" + input.AttemptResult.VisibilityRecoveryOutcome.PaneVisibleSource.ToString()
                + ",visibilityRecoveryReason=" + input.AttemptResult.VisibilityRecoveryOutcome.Reason
                + ",visibilityRecoveryDegradedReason=" + input.AttemptResult.VisibilityRecoveryOutcome.DegradedReason
                + ",refreshSourceStatus=" + input.AttemptResult.RefreshSourceSelectionOutcome.Status.ToString()
                + ",refreshSourceSelectedSource=" + input.AttemptResult.RefreshSourceSelectionOutcome.SelectedSource.ToString()
                + ",refreshSourceSelectionReason=" + input.AttemptResult.RefreshSourceSelectionOutcome.SelectionReason
                + ",refreshSourceFallbackReasons=" + input.AttemptResult.RefreshSourceSelectionOutcome.FallbackReasons
                + ",refreshSourceCacheFallback=" + input.AttemptResult.RefreshSourceSelectionOutcome.IsCacheFallback.ToString()
                + ",refreshSourceRebuildRequired=" + input.AttemptResult.RefreshSourceSelectionOutcome.IsRebuildRequired.ToString()
                + ",refreshSourceCanContinue=" + input.AttemptResult.RefreshSourceSelectionOutcome.CanContinueRefresh.ToString()
                + ",refreshSourceFailureReason=" + input.AttemptResult.RefreshSourceSelectionOutcome.FailureReason
                + ",refreshSourceDegradedReason=" + input.AttemptResult.RefreshSourceSelectionOutcome.DegradedReason
                + ",rebuildFallbackStatus=" + input.AttemptResult.RebuildFallbackOutcome.Status.ToString()
                + ",rebuildFallbackRequired=" + input.AttemptResult.RebuildFallbackOutcome.IsRequired.ToString()
                + ",rebuildFallbackCanContinue=" + input.AttemptResult.RebuildFallbackOutcome.CanContinueRefresh.ToString()
                + ",rebuildFallbackSnapshotSource=" + input.AttemptResult.RebuildFallbackOutcome.SnapshotSource.ToString()
                + ",rebuildFallbackReasons=" + input.AttemptResult.RebuildFallbackOutcome.FallbackReasons
                + ",rebuildFallbackFailureReason=" + input.AttemptResult.RebuildFallbackOutcome.FailureReason
                + ",rebuildFallbackDegradedReason=" + input.AttemptResult.RebuildFallbackOutcome.DegradedReason
                + ",refreshCompleted=" + input.AttemptResult.IsRefreshCompleted.ToString()
                + ",foregroundGuaranteeTerminal=" + input.AttemptResult.IsForegroundGuaranteeTerminal.ToString()
                + ",foregroundGuaranteeRequired=" + input.AttemptResult.WasForegroundGuaranteeRequired.ToString()
                + ",foregroundGuaranteeStatus=" + input.AttemptResult.ForegroundGuaranteeOutcome.Status.ToString()
                + ",foregroundGuaranteeDisplayCompletable=" + input.AttemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable.ToString()
                + ",foregroundGuaranteeExecutionAttempted=" + input.AttemptResult.ForegroundGuaranteeOutcome.WasExecutionAttempted.ToString()
                + ",foregroundGuaranteeTargetKind=" + input.AttemptResult.ForegroundGuaranteeOutcome.TargetKind.ToString()
                + ",foregroundRecoverySucceeded="
                + (input.AttemptResult.ForegroundGuaranteeOutcome.RecoverySucceeded.HasValue
                    ? input.AttemptResult.ForegroundGuaranteeOutcome.RecoverySucceeded.Value.ToString()
                    : string.Empty)
                + ",foregroundOutcomeReason=" + input.AttemptResult.ForegroundGuaranteeOutcome.Reason
                + WindowActivateDownstreamObservation.FormatDisplayRequestTraceFields(input.DisplayRequest);
            if (input.AttemptNumber.HasValue)
            {
                details += ",attempt=" + input.AttemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            string kernelTraceMessage =
                KernelFlickerTracePrefix
                + " source=" + OrchestrationSource
                + " action=case-display-completed sessionId="
                + input.SessionId
                + ", reason="
                + (input.Reason ?? string.Empty)
                + ", workbook="
                + input.FormattedWorkbook
                + ", window="
                + input.FormattedWindow
                + ", completion="
                + input.AttemptResult.CompletionBasis;

            return new CaseDisplayCompletedPayload(
                kernelTraceMessage,
                "case-display-completed",
                CompleteCreatedCaseDisplaySessionSource,
                input.WorkbookFullName,
                details);
        }
    }

    internal sealed class CaseDisplayCompletedPayloadInput
    {
        internal CaseDisplayCompletedPayloadInput(
            string reason,
            string sessionId,
            string workbookFullName,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            TaskPaneDisplayRequest displayRequest,
            string formattedWorkbook,
            string formattedWindow)
        {
            Reason = reason;
            SessionId = sessionId ?? string.Empty;
            WorkbookFullName = workbookFullName ?? string.Empty;
            AttemptResult = attemptResult;
            CompletionSource = completionSource;
            AttemptNumber = attemptNumber;
            DisplayRequest = displayRequest;
            FormattedWorkbook = formattedWorkbook ?? string.Empty;
            FormattedWindow = formattedWindow ?? string.Empty;
        }

        internal string Reason { get; }

        internal string SessionId { get; }

        internal string WorkbookFullName { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string CompletionSource { get; }

        internal int? AttemptNumber { get; }

        internal TaskPaneDisplayRequest DisplayRequest { get; }

        internal string FormattedWorkbook { get; }

        internal string FormattedWindow { get; }
    }

    internal sealed class CaseDisplayCompletedPayload
    {
        internal CaseDisplayCompletedPayload(
            string kernelTraceMessage,
            string observationAction,
            string observationSource,
            string workbookFullName,
            string details)
        {
            KernelTraceMessage = kernelTraceMessage ?? string.Empty;
            ObservationAction = observationAction ?? string.Empty;
            ObservationSource = observationSource ?? string.Empty;
            WorkbookFullName = workbookFullName ?? string.Empty;
            Details = details ?? string.Empty;
        }

        internal string KernelTraceMessage { get; }

        internal string ObservationAction { get; }

        internal string ObservationSource { get; }

        internal string WorkbookFullName { get; }

        internal string Details { get; }
    }
}
