namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshAttemptResult
    {
        private TaskPaneRefreshAttemptResult(
            bool isRefreshSucceeded,
            bool wasSkipped = false,
            bool wasContextRejected = false,
            bool isPaneVisible = false,
            bool isRefreshCompleted = false,
            bool isForegroundGuaranteeTerminal = false,
            bool wasForegroundGuaranteeRequired = false,
            string completionBasis = "")
        {
            IsRefreshSucceeded = isRefreshSucceeded;
            WasSkipped = wasSkipped;
            WasContextRejected = wasContextRejected;
            IsPaneVisible = isPaneVisible;
            IsRefreshCompleted = isRefreshCompleted;
            IsForegroundGuaranteeTerminal = isForegroundGuaranteeTerminal;
            WasForegroundGuaranteeRequired = wasForegroundGuaranteeRequired;
            CompletionBasis = completionBasis ?? string.Empty;
        }

        internal static TaskPaneRefreshAttemptResult Succeeded()
        {
            return Succeeded(
                foregroundGuaranteeRequired: false,
                completionBasis: "refreshCompleted");
        }

        internal static TaskPaneRefreshAttemptResult Succeeded(bool foregroundGuaranteeRequired, string completionBasis)
        {
            return new TaskPaneRefreshAttemptResult(
                true,
                isPaneVisible: true,
                isRefreshCompleted: true,
                isForegroundGuaranteeTerminal: true,
                wasForegroundGuaranteeRequired: foregroundGuaranteeRequired,
                completionBasis: completionBasis);
        }

        internal static TaskPaneRefreshAttemptResult VisibleAlreadySatisfied()
        {
            return new TaskPaneRefreshAttemptResult(
                true,
                isPaneVisible: true,
                isRefreshCompleted: false,
                isForegroundGuaranteeTerminal: true,
                wasForegroundGuaranteeRequired: false,
                completionBasis: "visibleCasePaneAlreadyShown");
        }

        internal static TaskPaneRefreshAttemptResult Failed()
        {
            return new TaskPaneRefreshAttemptResult(false, completionBasis: "failed");
        }

        internal static TaskPaneRefreshAttemptResult Skipped()
        {
            return new TaskPaneRefreshAttemptResult(false, wasSkipped: true, completionBasis: "skipped");
        }

        internal static TaskPaneRefreshAttemptResult ContextRejected()
        {
            return new TaskPaneRefreshAttemptResult(false, wasContextRejected: true, completionBasis: "contextRejected");
        }

        internal bool IsRefreshSucceeded { get; }

        internal bool WasSkipped { get; }

        internal bool WasContextRejected { get; }

        internal bool IsPaneVisible { get; }

        internal bool IsRefreshCompleted { get; }

        internal bool IsForegroundGuaranteeTerminal { get; }

        internal bool WasForegroundGuaranteeRequired { get; }

        internal string CompletionBasis { get; }
    }
}
