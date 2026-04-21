using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneManagerOrchestrationPolicyTests
    {
        [Fact]
        public void ShouldHideAllAndSkip_ReturnsTrue_WhenContextIsMissing()
        {
            bool result = TaskPaneRefreshPreconditionPolicy.ShouldHideAllAndSkip(WorkbookRole.Unknown, windowKey: null);

            Assert.True(result);
        }

        [Fact]
        public void ShouldHideAllAndSkip_ReturnsTrue_WhenWorkbookRoleIsUnknown()
        {
            bool result = TaskPaneRefreshPreconditionPolicy.ShouldHideAllAndSkip(WorkbookRole.Unknown, windowKey: null);

            Assert.True(result);
        }

        [Fact]
        public void ShouldHideAllAndSkip_ReturnsTrue_WhenWindowKeyIsBlank()
        {
            bool result = TaskPaneRefreshPreconditionPolicy.ShouldHideAllAndSkip(WorkbookRole.Case, windowKey: " ");

            Assert.True(result);
        }

        [Fact]
        public void ShouldHideAllAndSkip_ReturnsFalse_WhenContextAndWindowKeyAreValid()
        {
            bool result = TaskPaneRefreshPreconditionPolicy.ShouldHideAllAndSkip(WorkbookRole.Kernel, windowKey: "100");

            Assert.False(result);
        }

        [Fact]
        public void Decide_ReturnsNone_WhenKernelCaseCreationFlowIsInactive()
        {
            TaskPaneHostPreparationAction action = TaskPaneHostPreparationPolicy.Decide(
                isKernelCaseCreationFlowActive: false,
                isCaseHost: true);

            Assert.Equal(TaskPaneHostPreparationAction.None, action);
        }

        [Fact]
        public void Decide_ReturnsHideNonCaseHostsExceptActiveWindow_ForCaseHostDuringCaseCreation()
        {
            TaskPaneHostPreparationAction action = TaskPaneHostPreparationPolicy.Decide(
                isKernelCaseCreationFlowActive: true,
                isCaseHost: true);

            Assert.Equal(TaskPaneHostPreparationAction.HideNonCaseHostsExceptActiveWindow, action);
        }

        [Fact]
        public void Decide_ReturnsHideAllExceptActiveWindow_ForNonCaseHostDuringCaseCreation()
        {
            TaskPaneHostPreparationAction action = TaskPaneHostPreparationPolicy.Decide(
                isKernelCaseCreationFlowActive: true,
                isCaseHost: false);

            Assert.Equal(TaskPaneHostPreparationAction.HideAllExceptActiveWindow, action);
        }

        [Theory]
        [InlineData(true, "WorkbookOpen", true)]
        [InlineData(true, "WorkbookActivate", true)]
        [InlineData(true, "SheetActivate", false)]
        [InlineData(false, "WorkbookOpen", false)]
        public void ShouldNotify_UsesSnapshotCacheUpdateAndLifecycleReason(
            bool updatedCaseSnapshotCache,
            string reason,
            bool expected)
        {
            bool result = CasePaneCacheRefreshNotificationPolicy.ShouldNotify(updatedCaseSnapshotCache, reason);

            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData(true, false, false)]
        [InlineData(false, true, false)]
        [InlineData(false, false, true)]
        public void ShouldReject_OnlyReturnsTrue_WhenNeitherExistingNorRenderPathApplies(
            bool showedExistingPane,
            bool shouldShowWithRenderPane,
            bool expected)
        {
            bool result = TaskPaneDisplayRejectPolicy.ShouldReject(showedExistingPane, shouldShowWithRenderPane);

            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData(true, false, 0)]
        [InlineData(false, true, 1)]
        [InlineData(false, false, 2)]
        public void Decide_ReturnsThreeWayDisplayResult(
            bool showedExistingPane,
            bool shouldShowWithRenderPane,
            int expected)
        {
            PaneDisplayPolicyResult result = PaneDisplayPolicy.Decide(showedExistingPane, shouldShowWithRenderPane);

            Assert.Equal((PaneDisplayPolicyResult)expected, result);
        }
    }
}
