namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshCompletionDecisionService
    {
        internal CreatedCaseDisplayCompletionDecision DecideCreatedCaseDisplayCompletion(
            TaskPaneRefreshCompletionDecisionInput input)
        {
            if (input == null || !input.IsCreatedCaseDisplayReason)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("reasonNotCreatedCaseDisplay");
            }

            TaskPaneRefreshAttemptResult attemptResult = input.AttemptResult;
            if (attemptResult == null)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("attemptResult=null");
            }

            if (!attemptResult.IsRefreshSucceeded)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("refreshSucceeded=false");
            }

            if (!attemptResult.IsPaneVisible)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("paneVisible=false");
            }

            if (attemptResult.VisibilityRecoveryOutcome == null)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("visibilityRecoveryOutcome=null");
            }

            if (!attemptResult.VisibilityRecoveryOutcome.IsTerminal)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("visibilityRecoveryTerminal=false");
            }

            if (!attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("visibilityRecoveryDisplayCompletable=false");
            }

            if (!IsForegroundDisplayCompletableTerminalInput(attemptResult.ForegroundGuaranteeOutcome))
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("foregroundGuaranteeDisplayCompletable=false");
            }

            return CreatedCaseDisplayCompletionDecision.Allowed();
        }

        internal static bool IsForegroundDisplayCompletableTerminalInput(ForegroundGuaranteeOutcome outcome)
        {
            return outcome != null
                && outcome.IsTerminal
                && outcome.IsDisplayCompletable;
        }
    }

    internal sealed class TaskPaneRefreshCompletionDecisionInput
    {
        internal TaskPaneRefreshCompletionDecisionInput(
            bool isCreatedCaseDisplayReason,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            IsCreatedCaseDisplayReason = isCreatedCaseDisplayReason;
            AttemptResult = attemptResult;
        }

        internal bool IsCreatedCaseDisplayReason { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }
    }

    internal struct CreatedCaseDisplayCompletionDecision
    {
        private CreatedCaseDisplayCompletionDecision(bool canComplete, string blockedReason)
        {
            CanComplete = canComplete;
            BlockedReason = blockedReason ?? string.Empty;
        }

        internal bool CanComplete { get; }

        internal string BlockedReason { get; }

        internal static CreatedCaseDisplayCompletionDecision Allowed()
        {
            return new CreatedCaseDisplayCompletionDecision(true, string.Empty);
        }

        internal static CreatedCaseDisplayCompletionDecision Blocked(string blockedReason)
        {
            return new CreatedCaseDisplayCompletionDecision(false, blockedReason);
        }
    }
}
