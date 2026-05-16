using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRefreshFailClosedOutcomeTests
    {
        [Fact]
        public void FromPreconditionDecision_PreservesWorkbookOpenSkipActionAsSkippedAttemptResult()
        {
            TaskPaneRefreshFailClosedOutcome outcome =
                TaskPaneRefreshFailClosedOutcome.FromPreconditionDecision(
                    TaskPaneRefreshPreconditionDecision.SkipWorkbookOpenWindowDependentRefresh());

            Assert.Equal("skip-workbook-open-window-dependent-refresh", outcome.SkipActionName);
            Assert.NotNull(outcome.AttemptResult);
            Assert.False(outcome.AttemptResult.IsRefreshSucceeded);
            Assert.True(outcome.AttemptResult.WasSkipped);
            Assert.Equal(outcome.SkipActionName, outcome.AttemptResult.CompletionBasis);
            Assert.Equal(
                ForegroundGuaranteeOutcomeStatus.SkippedNoKnownTarget,
                outcome.AttemptResult.ForegroundGuaranteeOutcome.Status);
        }

        [Fact]
        public void FromPreconditionDecision_PreservesProtectionSkipActionAsSkippedAttemptResult()
        {
            TaskPaneRefreshFailClosedOutcome outcome =
                TaskPaneRefreshFailClosedOutcome.FromPreconditionDecision(
                    TaskPaneRefreshPreconditionDecision.IgnoreDuringProtection());

            Assert.Equal("ignore-during-protection", outcome.SkipActionName);
            Assert.False(outcome.AttemptResult.IsRefreshSucceeded);
            Assert.True(outcome.AttemptResult.WasSkipped);
            Assert.Equal(outcome.SkipActionName, outcome.AttemptResult.CompletionBasis);
        }
    }
}
