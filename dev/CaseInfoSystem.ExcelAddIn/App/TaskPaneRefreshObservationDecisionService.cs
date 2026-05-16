using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshObservationDecisionService
    {
        internal TaskPaneRefreshObservationDecision CompleteNormalizedOutcomeChain(
            TaskPaneRefreshObservationDecisionInput input)
        {
            if (input == null || input.AttemptResult == null)
            {
                return TaskPaneRefreshObservationDecision.Empty();
            }

            TaskPaneRefreshAttemptResult currentAttemptResult = input.AttemptResult;
            VisibilityRecoveryOutcome visibilityOutcome = TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(
                input.Workbook,
                input.InputWindow,
                currentAttemptResult,
                input.WorkbookWindowEnsureFacts);
            currentAttemptResult = currentAttemptResult.WithVisibilityRecoveryOutcome(visibilityOutcome);
            TaskPaneRefreshVisibilityObservationDecision visibilityDecision =
                TaskPaneRefreshVisibilityObservationDecision.Create(
                    input,
                    currentAttemptResult,
                    visibilityOutcome);

            RefreshSourceSelectionOutcome refreshSourceOutcome =
                TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(currentAttemptResult);
            currentAttemptResult = currentAttemptResult.WithRefreshSourceSelectionOutcome(refreshSourceOutcome);
            TaskPaneRefreshSourceObservationDecision refreshSourceDecision =
                TaskPaneRefreshSourceObservationDecision.Create(
                    input,
                    currentAttemptResult,
                    refreshSourceOutcome);

            RebuildFallbackOutcome rebuildFallbackOutcome =
                TaskPaneNormalizedOutcomeMapper.BuildRebuildFallbackOutcome(currentAttemptResult);
            currentAttemptResult = currentAttemptResult.WithRebuildFallbackOutcome(rebuildFallbackOutcome);
            TaskPaneRefreshRebuildFallbackObservationDecision rebuildFallbackDecision =
                TaskPaneRefreshRebuildFallbackObservationDecision.Create(
                    input,
                    currentAttemptResult,
                    rebuildFallbackOutcome);

            return new TaskPaneRefreshObservationDecision(
                currentAttemptResult,
                visibilityDecision,
                refreshSourceDecision,
                rebuildFallbackDecision);
        }

        internal TaskPaneRefreshForegroundGuaranteeDecision DecideForegroundGuarantee(
            TaskPaneRefreshAttemptResult attemptResult,
            Excel.Window inputWindow)
        {
            if (attemptResult == null)
            {
                return TaskPaneRefreshForegroundGuaranteeDecision.NoExecution(
                    attemptResult,
                    ForegroundGuaranteeOutcome.Unknown("attemptResult=null"),
                    inputWindow,
                    foregroundSkipReason: "attemptResult=null");
            }

            ForegroundGuaranteeOutcome existingOutcome = attemptResult.ForegroundGuaranteeOutcome;
            if (existingOutcome != null
                && existingOutcome.Status == ForegroundGuaranteeOutcomeStatus.SkippedAlreadyVisible)
            {
                return TaskPaneRefreshForegroundGuaranteeDecision.NoExecution(
                    attemptResult,
                    existingOutcome,
                    inputWindow,
                    existingOutcome.Reason);
            }

            if (!attemptResult.IsRefreshSucceeded || !attemptResult.IsPaneVisible)
            {
                ForegroundGuaranteeOutcome skippedOutcome = attemptResult.WasSkipped || attemptResult.WasContextRejected
                    ? attemptResult.ForegroundGuaranteeOutcome
                    : ForegroundGuaranteeOutcome.NotRequired("refreshSucceeded=false");
                return TaskPaneRefreshForegroundGuaranteeDecision.NoExecution(
                    attemptResult.WithForegroundGuaranteeOutcome(skippedOutcome),
                    skippedOutcome,
                    inputWindow,
                    skippedOutcome == null ? string.Empty : skippedOutcome.Reason);
            }

            bool foregroundRecoveryStarted = attemptResult.IsRefreshCompleted
                && attemptResult.ForegroundWindow != null
                && attemptResult.IsForegroundRecoveryServiceAvailable;
            string foregroundSkipReason = ResolveForegroundSkipReason(attemptResult);
            if (!foregroundRecoveryStarted)
            {
                ForegroundGuaranteeOutcome notRequiredOutcome = ForegroundGuaranteeOutcome.NotRequired(foregroundSkipReason);
                return TaskPaneRefreshForegroundGuaranteeDecision.NoExecution(
                    attemptResult.WithForegroundGuaranteeOutcome(notRequiredOutcome),
                    notRequiredOutcome,
                    inputWindow,
                    foregroundSkipReason);
            }

            ForegroundGuaranteeTargetKind targetKind = attemptResult.ForegroundWorkbook == null
                ? ForegroundGuaranteeTargetKind.ActiveWorkbookFallback
                : ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow;
            return TaskPaneRefreshForegroundGuaranteeDecision.Execute(
                attemptResult,
                ForegroundGuaranteeOutcome.Unknown("executionPending"),
                inputWindow,
                targetKind);
        }

        internal ForegroundGuaranteeOutcome ClassifyRequiredForegroundExecutionOutcome(
            ForegroundGuaranteeTargetKind targetKind,
            ForegroundGuaranteeExecutionResult executionResult)
        {
            return executionResult.ExecutionAttempted && executionResult.Recovered
                ? ForegroundGuaranteeOutcome.RequiredSucceeded(targetKind, "foregroundRecoverySucceeded")
                : ForegroundGuaranteeOutcome.RequiredDegraded(targetKind, "foregroundRecoveryReturnedFalse");
        }

        private static string ResolveForegroundSkipReason(TaskPaneRefreshAttemptResult attemptResult)
        {
            if (!attemptResult.IsRefreshCompleted)
            {
                return "refreshCompleted=false";
            }

            if (attemptResult.ForegroundWindow == null)
            {
                return "window=null";
            }

            if (!attemptResult.IsForegroundRecoveryServiceAvailable)
            {
                return "recoveryService=null";
            }

            return string.Empty;
        }
    }

    internal sealed class TaskPaneRefreshObservationDecisionInput
    {
        internal TaskPaneRefreshObservationDecisionInput(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            Reason = reason;
            Workbook = workbook;
            InputWindow = inputWindow;
            AttemptResult = attemptResult;
            CompletionSource = completionSource;
            AttemptNumber = attemptNumber;
            WorkbookWindowEnsureFacts = workbookWindowEnsureFacts;
        }

        internal string Reason { get; }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window InputWindow { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string CompletionSource { get; }

        internal int? AttemptNumber { get; }

        internal WorkbookWindowVisibilityEnsureFacts WorkbookWindowEnsureFacts { get; }
    }

    internal sealed class TaskPaneRefreshObservationDecision
    {
        internal TaskPaneRefreshObservationDecision(
            TaskPaneRefreshAttemptResult attemptResult,
            TaskPaneRefreshVisibilityObservationDecision visibility,
            TaskPaneRefreshSourceObservationDecision refreshSource,
            TaskPaneRefreshRebuildFallbackObservationDecision rebuildFallback)
        {
            AttemptResult = attemptResult;
            Visibility = visibility;
            RefreshSource = refreshSource;
            RebuildFallback = rebuildFallback;
        }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal TaskPaneRefreshVisibilityObservationDecision Visibility { get; }

        internal TaskPaneRefreshSourceObservationDecision RefreshSource { get; }

        internal TaskPaneRefreshRebuildFallbackObservationDecision RebuildFallback { get; }

        internal static TaskPaneRefreshObservationDecision Empty()
        {
            return new TaskPaneRefreshObservationDecision(null, null, null, null);
        }
    }

    internal sealed class TaskPaneRefreshVisibilityObservationDecision
    {
        private TaskPaneRefreshVisibilityObservationDecision(
            TaskPaneRefreshAttemptResult attemptResult,
            VisibilityRecoveryOutcome outcome,
            WorkbookContext context,
            Excel.Window observedWindow,
            string details)
        {
            AttemptResult = attemptResult;
            Outcome = outcome;
            Context = context;
            ObservedWindow = observedWindow;
            Details = details ?? string.Empty;
        }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal VisibilityRecoveryOutcome Outcome { get; }

        internal WorkbookContext Context { get; }

        internal Excel.Window ObservedWindow { get; }

        internal string Details { get; }

        internal static TaskPaneRefreshVisibilityObservationDecision Create(
            TaskPaneRefreshObservationDecisionInput input,
            TaskPaneRefreshAttemptResult attemptResult,
            VisibilityRecoveryOutcome outcome)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            return new TaskPaneRefreshVisibilityObservationDecision(
                attemptResult,
                outcome,
                context,
                context == null ? input.InputWindow : context.Window,
                TaskPaneNormalizedOutcomeMapper.FormatVisibilityRecoveryDetails(
                    input.Reason,
                    outcome,
                    attemptResult,
                    input.CompletionSource,
                    input.AttemptNumber,
                    input.WorkbookWindowEnsureFacts));
        }
    }

    internal sealed class TaskPaneRefreshSourceObservationDecision
    {
        private TaskPaneRefreshSourceObservationDecision(
            TaskPaneRefreshAttemptResult attemptResult,
            RefreshSourceSelectionOutcome outcome,
            WorkbookContext context,
            Excel.Window observedWindow,
            string statusAction,
            bool shouldLogRebuildRequiredTrace,
            string details)
        {
            AttemptResult = attemptResult;
            Outcome = outcome;
            Context = context;
            ObservedWindow = observedWindow;
            StatusAction = statusAction ?? string.Empty;
            ShouldLogRebuildRequiredTrace = shouldLogRebuildRequiredTrace;
            Details = details ?? string.Empty;
        }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal RefreshSourceSelectionOutcome Outcome { get; }

        internal WorkbookContext Context { get; }

        internal Excel.Window ObservedWindow { get; }

        internal string StatusAction { get; }

        internal bool ShouldLogRebuildRequiredTrace { get; }

        internal string Details { get; }

        internal static TaskPaneRefreshSourceObservationDecision Create(
            TaskPaneRefreshObservationDecisionInput input,
            TaskPaneRefreshAttemptResult attemptResult,
            RefreshSourceSelectionOutcome outcome)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            string details = TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionDetails(
                input.Reason,
                outcome,
                attemptResult,
                input.CompletionSource,
                input.AttemptNumber);
            string statusAction = TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(outcome);
            return new TaskPaneRefreshSourceObservationDecision(
                attemptResult,
                outcome,
                context,
                context == null ? input.InputWindow : context.Window,
                statusAction,
                outcome.IsRebuildRequired && !string.Equals(statusAction, "refresh-source-rebuild-required", System.StringComparison.OrdinalIgnoreCase),
                details);
        }
    }

    internal sealed class TaskPaneRefreshRebuildFallbackObservationDecision
    {
        private TaskPaneRefreshRebuildFallbackObservationDecision(
            TaskPaneRefreshAttemptResult attemptResult,
            RebuildFallbackOutcome outcome,
            WorkbookContext context,
            Excel.Window observedWindow,
            string statusAction,
            string details)
        {
            AttemptResult = attemptResult;
            Outcome = outcome;
            Context = context;
            ObservedWindow = observedWindow;
            StatusAction = statusAction ?? string.Empty;
            Details = details ?? string.Empty;
        }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal RebuildFallbackOutcome Outcome { get; }

        internal WorkbookContext Context { get; }

        internal Excel.Window ObservedWindow { get; }

        internal string StatusAction { get; }

        internal string Details { get; }

        internal static TaskPaneRefreshRebuildFallbackObservationDecision Create(
            TaskPaneRefreshObservationDecisionInput input,
            TaskPaneRefreshAttemptResult attemptResult,
            RebuildFallbackOutcome outcome)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            string details = TaskPaneNormalizedOutcomeMapper.FormatRebuildFallbackDetails(
                input.Reason,
                outcome,
                attemptResult,
                input.CompletionSource,
                input.AttemptNumber);
            string statusAction = "rebuild-fallback-" + outcome.Status.ToString().ToLowerInvariant();
            return new TaskPaneRefreshRebuildFallbackObservationDecision(
                attemptResult,
                outcome,
                context,
                context == null ? input.InputWindow : context.Window,
                statusAction,
                details);
        }
    }

    internal sealed class TaskPaneRefreshForegroundGuaranteeDecision
    {
        private TaskPaneRefreshForegroundGuaranteeDecision(
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeOutcome outcome,
            bool shouldExecuteForegroundGuarantee,
            ForegroundGuaranteeTargetKind targetKind,
            bool foregroundRecoveryStarted,
            string foregroundSkipReason,
            Excel.Window inputWindow)
        {
            AttemptResult = attemptResult;
            Outcome = outcome;
            ShouldExecuteForegroundGuarantee = shouldExecuteForegroundGuarantee;
            TargetKind = targetKind;
            ForegroundRecoveryStarted = foregroundRecoveryStarted;
            ForegroundSkipReason = foregroundSkipReason ?? string.Empty;
            Context = attemptResult == null ? null : attemptResult.ForegroundContext;
            ResolvedWindow = attemptResult == null || attemptResult.ForegroundWindow == null
                ? inputWindow
                : attemptResult.ForegroundWindow;
            ObservedWindow = Context == null ? inputWindow : Context.Window;
            RecoveryServicePresent = attemptResult != null && attemptResult.IsForegroundRecoveryServiceAvailable;
        }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal ForegroundGuaranteeOutcome Outcome { get; }

        internal bool ShouldExecuteForegroundGuarantee { get; }

        internal ForegroundGuaranteeTargetKind TargetKind { get; }

        internal bool ForegroundRecoveryStarted { get; }

        internal string ForegroundSkipReason { get; }

        internal WorkbookContext Context { get; }

        internal Excel.Window ResolvedWindow { get; }

        internal Excel.Window ObservedWindow { get; }

        internal bool RecoveryServicePresent { get; }

        internal static TaskPaneRefreshForegroundGuaranteeDecision NoExecution(
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeOutcome outcome,
            Excel.Window inputWindow,
            string foregroundSkipReason)
        {
            return new TaskPaneRefreshForegroundGuaranteeDecision(
                attemptResult,
                outcome,
                shouldExecuteForegroundGuarantee: false,
                targetKind: outcome == null ? ForegroundGuaranteeTargetKind.Unknown : outcome.TargetKind,
                foregroundRecoveryStarted: false,
                foregroundSkipReason: foregroundSkipReason,
                inputWindow: inputWindow);
        }

        internal static TaskPaneRefreshForegroundGuaranteeDecision Execute(
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeOutcome outcome,
            Excel.Window inputWindow,
            ForegroundGuaranteeTargetKind targetKind)
        {
            return new TaskPaneRefreshForegroundGuaranteeDecision(
                attemptResult,
                outcome,
                shouldExecuteForegroundGuarantee: true,
                targetKind: targetKind,
                foregroundRecoveryStarted: true,
                foregroundSkipReason: string.Empty,
                inputWindow: inputWindow);
        }
    }
}
