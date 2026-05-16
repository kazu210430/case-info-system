using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshRetryContinuationDecisionService
    {
        internal TaskPaneRefreshRetryContinuationDecision DecideBeforeTick(bool hasAttemptsRemaining)
        {
            return hasAttemptsRemaining
                ? TaskPaneRefreshRetryContinuationDecision.ContinueRetrySequence("attemptsRemaining")
                : TaskPaneRefreshRetryContinuationDecision.StopRetrySequence("attemptsExhausted");
        }

        internal TaskPaneRefreshRetryContinuationDecision DecideAfterWorkbookTargetResolution(bool hasTargetWorkbook)
        {
            return hasTargetWorkbook
                ? TaskPaneRefreshRetryContinuationDecision.ContinueRetrySequence("workbookTargetResolved")
                : TaskPaneRefreshRetryContinuationDecision.ContinueToActiveContextFallback("workbookTargetMissing");
        }

        internal TaskPaneRefreshRetryContinuationDecision DecideActiveContextFallback(WorkbookContext context)
        {
            return context != null && context.Role == WorkbookRole.Case
                ? TaskPaneRefreshRetryContinuationDecision.AttemptActiveContextFallback("activeCaseContext")
                : TaskPaneRefreshRetryContinuationDecision.StopRetrySequence(
                    context == null ? "activeContextMissing" : "activeContextNotCase");
        }

        internal TaskPaneRefreshRetryContinuationDecision DecideAfterRefresh(bool refreshed)
        {
            return refreshed
                ? TaskPaneRefreshRetryContinuationDecision.StopRetrySequence("refreshSucceeded")
                : TaskPaneRefreshRetryContinuationDecision.ContinueRetrySequence("refreshFailed");
        }
    }

    internal sealed class TaskPaneRefreshRetryContinuationDecision
    {
        private TaskPaneRefreshRetryContinuationDecision(
            bool handled,
            bool shouldStopTimer,
            bool shouldAttemptActiveContextFallback,
            string action,
            string description)
        {
            Handled = handled;
            ShouldStopTimer = shouldStopTimer;
            ShouldAttemptActiveContextFallback = shouldAttemptActiveContextFallback;
            Action = action ?? string.Empty;
            Description = description ?? string.Empty;
        }

        internal bool Handled { get; }

        internal bool ShouldStopTimer { get; }

        internal bool ShouldAttemptActiveContextFallback { get; }

        internal string Action { get; }

        internal string Description { get; }

        internal static TaskPaneRefreshRetryContinuationDecision ContinueToActiveContextFallback(string description)
        {
            return new TaskPaneRefreshRetryContinuationDecision(
                handled: false,
                shouldStopTimer: false,
                shouldAttemptActiveContextFallback: false,
                action: "continue-to-active-context-fallback",
                description: description);
        }

        internal static TaskPaneRefreshRetryContinuationDecision AttemptActiveContextFallback(string description)
        {
            return new TaskPaneRefreshRetryContinuationDecision(
                handled: true,
                shouldStopTimer: false,
                shouldAttemptActiveContextFallback: true,
                action: "attempt-active-context-fallback",
                description: description);
        }

        internal static TaskPaneRefreshRetryContinuationDecision ContinueRetrySequence(string description)
        {
            return new TaskPaneRefreshRetryContinuationDecision(
                handled: true,
                shouldStopTimer: false,
                shouldAttemptActiveContextFallback: false,
                action: "continue-retry",
                description: description);
        }

        internal static TaskPaneRefreshRetryContinuationDecision StopRetrySequence(string description)
        {
            return new TaskPaneRefreshRetryContinuationDecision(
                handled: true,
                shouldStopTimer: true,
                shouldAttemptActiveContextFallback: false,
                action: "stop-retry",
                description: description);
        }
    }
}
