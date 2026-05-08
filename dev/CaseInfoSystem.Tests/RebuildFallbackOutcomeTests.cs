using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class RebuildFallbackOutcomeTests
    {
        [Fact]
        public void FromBuildResult_WhenBaseCacheFallbackSuppliesSnapshot_SkipsRebuildFallback()
        {
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                "snapshot",
                updatedCaseSnapshotCache: true,
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.BaseCacheFallback,
                "LatestMasterVersionUnavailable",
                masterListRebuildAttempted: false,
                masterListRebuildSucceeded: false,
                failureReason: string.Empty,
                degradedReason: string.Empty);

            RebuildFallbackOutcome outcome = RebuildFallbackOutcome.FromBuildResult(buildResult);

            Assert.Equal(RebuildFallbackOutcomeStatus.Skipped, outcome.Status);
            Assert.False(outcome.IsRequired);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.CanContinueRefresh);
            Assert.Equal(TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.BaseCacheFallback, outcome.SnapshotSource);
            Assert.Equal("LatestMasterVersionUnavailable", outcome.FallbackReasons);
        }

        [Fact]
        public void FromBuildResult_WhenMasterListRebuildSucceeds_CompletesRequiredFallback()
        {
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                "snapshot",
                updatedCaseSnapshotCache: true,
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                "CaseCacheStale|BaseCacheStale",
                masterListRebuildAttempted: true,
                masterListRebuildSucceeded: true,
                failureReason: string.Empty,
                degradedReason: string.Empty);

            RebuildFallbackOutcome outcome = RebuildFallbackOutcome.FromBuildResult(buildResult);

            Assert.Equal(RebuildFallbackOutcomeStatus.Completed, outcome.Status);
            Assert.True(outcome.IsRequired);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.CanContinueRefresh);
            Assert.True(outcome.MasterListRebuildAttempted);
            Assert.True(outcome.MasterListRebuildSucceeded);
        }

        [Fact]
        public void FromBuildResult_WhenMasterListRebuildReturnsErrorSnapshot_MarksDegraded()
        {
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                "META\t2\tERROR\tstep=20",
                updatedCaseSnapshotCache: false,
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                "CacheUnavailable",
                masterListRebuildAttempted: true,
                masterListRebuildSucceeded: false,
                failureReason: "SnapshotBuildException",
                degradedReason: "SnapshotBuildException");

            RebuildFallbackOutcome outcome = RebuildFallbackOutcome.FromBuildResult(buildResult);

            Assert.Equal(RebuildFallbackOutcomeStatus.Degraded, outcome.Status);
            Assert.True(outcome.IsRequired);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.CanContinueRefresh);
            Assert.False(outcome.MasterListRebuildSucceeded);
            Assert.Equal("SnapshotBuildException", outcome.DegradedReason);
        }
    }
}
