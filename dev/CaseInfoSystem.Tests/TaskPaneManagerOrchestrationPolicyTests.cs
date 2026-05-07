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
        public void DecideHostFlowPrecondition_ReturnsHideAllAndSkipForUnknownRole()
        {
            TaskPaneHostFlowPreconditionDecision result = TaskPaneRefreshPreconditionPolicy.DecideHostFlowPrecondition(
                WorkbookRole.Unknown,
                windowKey: null);

            Assert.Equal(TaskPaneHostFlowPreconditionDecision.HideAllAndSkipForUnknownRole, result);
        }

        [Fact]
        public void DecideHostFlowPrecondition_ReturnsHideAllAndSkipForMissingWindowKey()
        {
            TaskPaneHostFlowPreconditionDecision result = TaskPaneRefreshPreconditionPolicy.DecideHostFlowPrecondition(
                WorkbookRole.Case,
                windowKey: " ");

            Assert.Equal(TaskPaneHostFlowPreconditionDecision.HideAllAndSkipForMissingWindowKey, result);
        }

        [Fact]
        public void DecideHostFlowPrecondition_ReturnsProceed_WhenRoleAndWindowKeyAreValid()
        {
            TaskPaneHostFlowPreconditionDecision result = TaskPaneRefreshPreconditionPolicy.DecideHostFlowPrecondition(
                WorkbookRole.Kernel,
                windowKey: "100");

            Assert.Equal(TaskPaneHostFlowPreconditionDecision.Proceed, result);
        }

        [Fact]
        public void DecideHostFlowPrecondition_ReturnsProceed_WhenWindowKeyIsNotEvaluatedYet()
        {
            TaskPaneHostFlowPreconditionDecision result = TaskPaneRefreshPreconditionPolicy.DecideHostFlowPrecondition(
                WorkbookRole.Case,
                windowKey: null);

            Assert.Equal(TaskPaneHostFlowPreconditionDecision.Proceed, result);
        }

        [Fact]
        public void DecideHostFlowPrecondition_PrioritizesUnknownRole_OverMissingWindowKey()
        {
            TaskPaneHostFlowPreconditionDecision result = TaskPaneRefreshPreconditionPolicy.DecideHostFlowPrecondition(
                WorkbookRole.Unknown,
                windowKey: " ");

            Assert.Equal(TaskPaneHostFlowPreconditionDecision.HideAllAndSkipForUnknownRole, result);
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
        [InlineData("WorkbookOpen", true, false, true)]
        [InlineData("WorkbookOpen", true, true, false)]
        [InlineData("WorkbookActivate", true, false, false)]
        [InlineData("WorkbookOpen", false, false, false)]
        public void ShouldSkip_UsesWorkbookOpenWindowDependencyBoundary(
            string reason,
            bool hasWorkbook,
            bool hasWindow,
            bool expected)
        {
            Excel.Workbook workbook = hasWorkbook ? new Excel.Workbook() : null;
            Excel.Window window = hasWindow ? new Excel.Window { Hwnd = 101 } : null;

            bool result = TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window);

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
            TaskPaneDisplayEntryDecision decision = PaneDisplayPolicy.Decide(
                request: null,
                taskPaneManager: null,
                workbookRoleResolver: null,
                workbook: null,
                window: new Excel.Window { Hwnd = 101 });

            Assert.Equal(PaneDisplayPolicyResult.Reject, decision.Result);
        }

        [Fact]
        public void Decide_ReturnsReject_WhenTargetWindowIsMissing()
        {
            TaskPaneDisplayEntryDecision decision = PaneDisplayPolicy.Decide(
                TaskPaneDisplayRequest.ForWindowActivate(),
                taskPaneManager: null,
                workbookRoleResolver: null,
                workbook: null,
                window: null);

            Assert.Equal(PaneDisplayPolicyResult.Reject, decision.Result);
        }

        [Fact]
        public void Decide_ReturnsHide_WhenRequestIsAccepted_AndManagedPaneRemains_ForNonDisplayTarget()
        {
            TaskPaneDisplayEntryDecision decision = PaneDisplayPolicy.Decide(
                TaskPaneDisplayRequest.ForWindowActivate(),
                new TaskPaneDisplayEntryState(
                    hasTargetWindow: true,
                    hasResolvableWindowKey: true,
                    hasManagedPane: true,
                    hasExistingHost: true,
                    isSameWorkbook: false,
                    isRenderSignatureCurrent: false),
                shouldDisplayPane: false);

            Assert.Equal(PaneDisplayPolicyResult.Hide, decision.Result);
        }

        [Fact]
        public void Decide_ReturnsShowExisting_WhenStateIsCurrentForSameWorkbook()
        {
            TaskPaneDisplayEntryState state = new TaskPaneDisplayEntryState(
                hasTargetWindow: true,
                hasResolvableWindowKey: true,
                hasManagedPane: true,
                hasExistingHost: true,
                isSameWorkbook: true,
                isRenderSignatureCurrent: true);

            TaskPaneDisplayEntryDecision decision = PaneDisplayPolicy.Decide(
                TaskPaneDisplayRequest.ForWindowActivate(),
                state,
                shouldDisplayPane: true);

            Assert.Equal(PaneDisplayPolicyResult.ShowExisting, decision.Result);
        }

        [Fact]
        public void Decide_ReturnsShowWithRender_WhenStateRequiresRerender()
        {
            TaskPaneDisplayEntryState state = new TaskPaneDisplayEntryState(
                hasTargetWindow: true,
                hasResolvableWindowKey: true,
                hasManagedPane: true,
                hasExistingHost: true,
                isSameWorkbook: true,
                isRenderSignatureCurrent: false);

            TaskPaneDisplayEntryDecision decision = PaneDisplayPolicy.Decide(
                TaskPaneDisplayRequest.ForWindowActivate(),
                state,
                shouldDisplayPane: true);

            Assert.Equal(PaneDisplayPolicyResult.ShowWithRender, decision.Result);
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
