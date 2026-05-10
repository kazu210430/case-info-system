using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class WorkbookCaseTaskPaneRefreshCommandNotificationPolicyTests
    {
        [Fact]
        public void Decide_ReturnsUpdated_WhenCaseSnapshotCacheWasUpdated()
        {
            TaskPaneRefreshAttemptResult result = CreateRefreshResult(
                new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                    "snapshot",
                    updatedCaseSnapshotCache: true));

            WorkbookCaseTaskPaneRefreshCommandNotificationKind kind =
                WorkbookCaseTaskPaneRefreshCommandNotificationPolicy.Decide(result);

            Assert.Equal(WorkbookCaseTaskPaneRefreshCommandNotificationKind.Updated, kind);
        }

        [Fact]
        public void Decide_ReturnsLatest_WhenRefreshSucceededWithoutCaseSnapshotCacheUpdate()
        {
            TaskPaneRefreshAttemptResult result = CreateRefreshResult(
                new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                    "snapshot",
                    updatedCaseSnapshotCache: false));

            WorkbookCaseTaskPaneRefreshCommandNotificationKind kind =
                WorkbookCaseTaskPaneRefreshCommandNotificationPolicy.Decide(result);

            Assert.Equal(WorkbookCaseTaskPaneRefreshCommandNotificationKind.Latest, kind);
        }

        [Fact]
        public void Decide_ReturnsLatest_WhenProtectionSkippedRibbonRefreshAsNoOp()
        {
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult.Skipped("ignore-during-protection");

            WorkbookCaseTaskPaneRefreshCommandNotificationKind kind =
                WorkbookCaseTaskPaneRefreshCommandNotificationPolicy.Decide(result);

            Assert.Equal(WorkbookCaseTaskPaneRefreshCommandNotificationKind.Latest, kind);
        }

        [Fact]
        public void Decide_ReturnsFailed_WhenRefreshFailed()
        {
            WorkbookCaseTaskPaneRefreshCommandNotificationKind kind =
                WorkbookCaseTaskPaneRefreshCommandNotificationPolicy.Decide(TaskPaneRefreshAttemptResult.Failed());

            Assert.Equal(WorkbookCaseTaskPaneRefreshCommandNotificationKind.Failed, kind);
        }

        private static TaskPaneRefreshAttemptResult CreateRefreshResult(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult)
        {
            return TaskPaneRefreshAttemptResult.RefreshCompletedPendingForeground(
                foregroundContext: null,
                foregroundWorkbook: null,
                foregroundWindow: null,
                isForegroundRecoveryServiceAvailable: false,
                completionBasis: "refreshCompleted",
                paneVisibleSource: PaneVisibleSource.RefreshedShown,
                snapshotBuildResult: buildResult,
                preContextRecoveryAttempted: false,
                preContextRecoverySucceeded: null);
        }
    }
}
