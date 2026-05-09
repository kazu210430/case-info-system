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

        [Fact]
        public void RefreshSourceOutcome_WhenCaseCacheSuppliesSnapshot_MarksSelected()
        {
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                "snapshot",
                updatedCaseSnapshotCache: false,
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.CaseCache,
                string.Empty,
                masterListRebuildAttempted: false,
                masterListRebuildSucceeded: false,
                failureReason: string.Empty,
                degradedReason: string.Empty);

            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(CreateShownAttempt(buildResult));

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.Selected, outcome.Status);
            Assert.Equal(TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.CaseCache, outcome.SelectedSource);
            Assert.False(outcome.IsCacheFallback);
            Assert.False(outcome.IsRebuildRequired);
            Assert.True(outcome.CanContinueRefresh);
        }

        [Fact]
        public void RefreshSourceOutcome_WhenBaseCacheFallbackSuppliesSnapshot_MarksFallbackSelected()
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

            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(CreateShownAttempt(buildResult));

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.FallbackSelected, outcome.Status);
            Assert.True(outcome.IsCacheFallback);
            Assert.False(outcome.IsRebuildRequired);
            Assert.Equal("LatestMasterVersionUnavailable", outcome.SelectionReason);
        }

        [Fact]
        public void RefreshSourceOutcome_WhenMasterListRebuildSucceeds_MarksRebuildRequired()
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

            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(CreateShownAttempt(buildResult));

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.RebuildRequired, outcome.Status);
            Assert.True(outcome.IsRebuildRequired);
            Assert.False(outcome.IsCacheFallback);
            Assert.True(outcome.CanContinueRefresh);
        }

        [Fact]
        public void RefreshSourceOutcome_WhenMasterListRebuildReturnsErrorSnapshot_MarksDegradedSelected()
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

            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(CreateShownAttempt(buildResult));

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.DegradedSelected, outcome.Status);
            Assert.True(outcome.IsRebuildRequired);
            Assert.True(outcome.CanContinueRefresh);
            Assert.Equal("SnapshotBuildException", outcome.DegradedReason);
        }

        [Fact]
        public void RefreshSourceOutcome_WhenSnapshotSelectionFails_MarksFailed()
        {
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                string.Empty,
                updatedCaseSnapshotCache: false,
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                "CacheUnavailable",
                masterListRebuildAttempted: true,
                masterListRebuildSucceeded: false,
                failureReason: "SnapshotBuildException",
                degradedReason: string.Empty);

            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(TaskPaneRefreshAttemptResult.Failed(buildResult));

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.Failed, outcome.Status);
            Assert.True(outcome.IsRebuildRequired);
            Assert.False(outcome.CanContinueRefresh);
            Assert.Equal("SnapshotBuildException", outcome.FailureReason);
        }

        [Fact]
        public void RefreshSourceOutcome_WhenSnapshotAcquisitionNotReached_MarksNotReached()
        {
            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied());

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.NotReached, outcome.Status);
            Assert.Equal(TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None, outcome.SelectedSource);
            Assert.False(outcome.IsRebuildRequired);
            Assert.True(outcome.CanContinueRefresh);
        }

        private static TaskPaneRefreshAttemptResult CreateShownAttempt(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult)
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

    public sealed class VisibilityForegroundOutcomeBoundaryTests
    {
        [Fact]
        public void VisibilityOutcome_WhenPaneVisibleWithDegradedRawFacts_RemainsVisibilityOutcomeOnly()
        {
            VisibilityRecoveryOutcome outcome = VisibilityRecoveryOutcome.Degraded(
                "paneVisibleWithDegradedRecoveryFacts",
                VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                PaneVisibleSource.RefreshedShown,
                WorkbookWindowVisibilityEnsureOutcome.MadeVisible,
                fullRecoveryAttempted: true,
                fullRecoverySucceeded: false,
                degradedReason: "preContextRecoveryReturnedFalse");

            Assert.Equal(VisibilityRecoveryOutcomeStatus.Degraded, outcome.Status);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsPaneVisible);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal(VisibilityRecoveryTargetKind.ExplicitWorkbookWindow, outcome.TargetKind);
            Assert.Equal(PaneVisibleSource.RefreshedShown, outcome.PaneVisibleSource);
            Assert.Equal(WorkbookWindowVisibilityEnsureOutcome.MadeVisible, outcome.WorkbookWindowEnsureStatus);
            Assert.True(outcome.FullRecoveryAttempted);
            Assert.False(outcome.FullRecoverySucceeded.GetValueOrDefault());
            Assert.Equal("preContextRecoveryReturnedFalse", outcome.DegradedReason);
        }

        [Fact]
        public void ForegroundOutcome_WhenRequiredExecutionReturnsFalse_RemainsForegroundOutcomeOnly()
        {
            ForegroundGuaranteeOutcome outcome = ForegroundGuaranteeOutcome.RequiredDegraded(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                "foregroundRecoveryReturnedFalse");

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredDegraded, outcome.Status);
            Assert.True(outcome.WasRequired);
            Assert.True(outcome.WasExecutionAttempted);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal(ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow, outcome.TargetKind);
            Assert.False(outcome.RecoverySucceeded.GetValueOrDefault());
            Assert.Equal("foregroundRecoveryReturnedFalse", outcome.Reason);
        }

        [Fact]
        public void RefreshAttemptResult_KeepsVisibilityAndForegroundOutcomesSeparateForCompletionGate()
        {
            VisibilityRecoveryOutcome visibilityOutcome = VisibilityRecoveryOutcome.Completed(
                "refreshedShown",
                VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                PaneVisibleSource.RefreshedShown,
                WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible,
                fullRecoveryAttempted: false,
                fullRecoverySucceeded: null);
            ForegroundGuaranteeOutcome foregroundOutcome = ForegroundGuaranteeOutcome.RequiredFailed(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                "foregroundRecoveryNotAttempted");

            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(visibilityOutcome)
                .WithForegroundGuaranteeOutcome(foregroundOutcome);

            Assert.True(result.IsRefreshSucceeded);
            Assert.True(result.IsPaneVisible);
            Assert.True(result.VisibilityRecoveryOutcome.IsDisplayCompletable);
            Assert.Equal(VisibilityRecoveryOutcomeStatus.Completed, result.VisibilityRecoveryOutcome.Status);
            Assert.True(result.IsForegroundGuaranteeTerminal);
            Assert.False(result.ForegroundGuaranteeOutcome.IsDisplayCompletable);
            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredFailed, result.ForegroundGuaranteeOutcome.Status);
        }
    }
}
