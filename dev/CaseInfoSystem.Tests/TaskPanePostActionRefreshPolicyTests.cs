using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class TaskPanePostActionRefreshPolicyTests
    {
        [Theory]
        [InlineData("doc")]
        [InlineData("DOC")]
        [InlineData("accounting")]
        public void Decide_ReturnsSkipForForegroundPreservation_ForForegroundSensitiveActions(string actionKind)
        {
            TaskPanePostActionRefreshDecision decision = TaskPanePostActionRefreshPolicy.Decide(actionKind);

            Assert.Equal(TaskPanePostActionRefreshDecision.SkipForForegroundPreservation, decision);
        }

        [Fact]
        public void Decide_ReturnsSkipForForegroundPreservation_ForCaseListAction()
        {
            TaskPanePostActionRefreshDecision decision = TaskPanePostActionRefreshPolicy.Decide("caselist");

            Assert.Equal(TaskPanePostActionRefreshDecision.SkipForForegroundPreservation, decision);
        }

        [Theory]
        [InlineData("preview")]
        [InlineData("")]
        [InlineData(null)]
        public void Decide_ReturnsRefreshImmediately_ForOtherActions(string actionKind)
        {
            TaskPanePostActionRefreshDecision decision = TaskPanePostActionRefreshPolicy.Decide(actionKind);

            Assert.Equal(TaskPanePostActionRefreshDecision.RefreshImmediately, decision);
        }
    }
}
