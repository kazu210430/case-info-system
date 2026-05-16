namespace CaseInfoSystem.ExcelAddIn.App
{
    internal readonly struct TaskPaneRefreshFailClosedOutcome
    {
        private TaskPaneRefreshFailClosedOutcome(TaskPaneRefreshAttemptResult attemptResult, string skipActionName)
        {
            AttemptResult = attemptResult;
            SkipActionName = skipActionName ?? string.Empty;
        }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string SkipActionName { get; }

        internal static TaskPaneRefreshFailClosedOutcome FromPreconditionDecision(
            TaskPaneRefreshPreconditionDecision preconditionDecision)
        {
            return new TaskPaneRefreshFailClosedOutcome(
                TaskPaneRefreshAttemptResult.Skipped(preconditionDecision.SkipActionName),
                preconditionDecision.SkipActionName);
        }
    }
}
