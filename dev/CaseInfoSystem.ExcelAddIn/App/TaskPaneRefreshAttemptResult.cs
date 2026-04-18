namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshAttemptResult
    {
        private TaskPaneRefreshAttemptResult(bool isRefreshSucceeded, bool wasSkipped = false, bool wasContextRejected = false)
        {
            IsRefreshSucceeded = isRefreshSucceeded;
            WasSkipped = wasSkipped;
            WasContextRejected = wasContextRejected;
        }

        internal static TaskPaneRefreshAttemptResult Succeeded()
        {
            return new TaskPaneRefreshAttemptResult(true);
        }

        internal static TaskPaneRefreshAttemptResult Failed()
        {
            return new TaskPaneRefreshAttemptResult(false);
        }

        internal static TaskPaneRefreshAttemptResult Skipped()
        {
            return new TaskPaneRefreshAttemptResult(false, wasSkipped: true);
        }

        internal static TaskPaneRefreshAttemptResult ContextRejected()
        {
            return new TaskPaneRefreshAttemptResult(false, wasContextRejected: true);
        }

        internal bool IsRefreshSucceeded { get; }

        internal bool WasSkipped { get; }

        internal bool WasContextRejected { get; }
    }
}
