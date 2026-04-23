using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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

        [Theory]
        [InlineData(false, true, false, false, 3)]
        [InlineData(false, false, false, false, 2)]
        [InlineData(true, true, false, true, 1)]
        public void Decide_UsesHideOnlyForNonDisplayTargetWithManagedPane(
            bool shouldDisplayPane,
            bool hasManagedPane,
            bool showedExistingPane,
            bool shouldShowWithRenderPane,
            int expected)
        {
            PaneDisplayPolicyResult result = PaneDisplayPolicy.Decide(
                shouldDisplayPane,
                hasManagedPane,
                showedExistingPane,
                shouldShowWithRenderPane);

            Assert.Equal((PaneDisplayPolicyResult)expected, result);
        }

        [Fact]
        public void Decide_ReturnsReject_WhenRequestIsNotAccepted()
        {
            PaneDisplayPolicyResult result = PaneDisplayPolicy.Decide(
                request: null,
                taskPaneManager: null,
                workbook: null,
                window: new Excel.Window { Hwnd = 101 },
                shouldDisplayPane: true);

            Assert.Equal(PaneDisplayPolicyResult.Reject, result);
        }

        [Fact]
        public void Decide_ReturnsReject_WhenTargetWindowIsMissing()
        {
            PaneDisplayPolicyResult result = PaneDisplayPolicy.Decide(
                TaskPaneDisplayRequest.ForWindowActivate(),
                taskPaneManager: null,
                workbook: null,
                window: null,
                shouldDisplayPane: true);

            Assert.Equal(PaneDisplayPolicyResult.Reject, result);
        }

        [Fact]
        public void Decide_ReturnsHide_WhenRequestIsAccepted_AndManagedPaneRemains_ForNonDisplayTarget()
        {
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new System.Collections.Generic.List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new System.Collections.Generic.List<string>()),
                testHooks: null);
            var targetWindow = new Excel.Window { Hwnd = 123 };
            manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new CaseInfoSystem.ExcelAddIn.UI.DocumentButtonsControl(), "123"));

            PaneDisplayPolicyResult result = PaneDisplayPolicy.Decide(
                TaskPaneDisplayRequest.ForWindowActivate(),
                manager,
                workbook: null,
                window: targetWindow,
                shouldDisplayPane: false);

            Assert.Equal(PaneDisplayPolicyResult.Hide, result);
        }

        [Theory]
        [InlineData(1, true)]
        [InlineData(2, true)]
        [InlineData(3, true)]
        [InlineData(0, false)]
        public void ShouldDisplayPane_UsesHandledWorkbookRoles(
            int role,
            bool expected)
        {
            var resolver = new FakeWorkbookRoleResolver
            {
                Role = (WorkbookRole)role
            };

            bool result = PaneDisplayPolicy.ShouldDisplayPane(resolver, workbook: null);

            Assert.Equal(expected, result);
        }

        [Fact]
        public void ShouldDisplayPane_ReturnsTrue_WhenResolverIsMissing()
        {
            bool result = PaneDisplayPolicy.ShouldDisplayPane(workbookRoleResolver: null, workbook: null);

            Assert.True(result);
        }
    }
}
