using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneNormalizedOutcomeMapperTests
    {
        [Fact]
        public void BuildVisibilityRecoveryOutcome_WhenRefreshFailed_RemainsNonDisplayCompletable()
        {
            VisibilityRecoveryOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(
                workbook: null,
                inputWindow: null,
                attemptResult: TaskPaneRefreshAttemptResult.Failed(),
                workbookWindowEnsureFacts: null);

            Assert.Equal(VisibilityRecoveryOutcomeStatus.Failed, outcome.Status);
            Assert.Equal("refreshFailed", outcome.Reason);
            Assert.True(outcome.IsTerminal);
            Assert.False(outcome.IsPaneVisible);
            Assert.False(outcome.IsDisplayCompletable);
            Assert.Equal(VisibilityRecoveryTargetKind.NoKnownTarget, outcome.TargetKind);
            Assert.Equal(PaneVisibleSource.None, outcome.PaneVisibleSource);
        }

        [Fact]
        public void BuildVisibilityRecoveryOutcome_WhenAlreadyVisible_RemainsDisplayCompletableSkippedOutcome()
        {
            VisibilityRecoveryOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(
                workbook: null,
                inputWindow: null,
                attemptResult: TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied(),
                workbookWindowEnsureFacts: null);

            Assert.Equal(VisibilityRecoveryOutcomeStatus.Skipped, outcome.Status);
            Assert.Equal("alreadyVisible", outcome.Reason);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsPaneVisible);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal(VisibilityRecoveryTargetKind.AlreadyVisible, outcome.TargetKind);
            Assert.Equal(PaneVisibleSource.AlreadyVisibleHost, outcome.PaneVisibleSource);
        }

        [Fact]
        public void BuildVisibilityRecoveryOutcome_WhenEnsureFactsAreDegraded_KeepsVisibilityDegradedDisplayCompletable()
        {
            WorkbookWindowVisibilityEnsureFacts ensureFacts = WorkbookWindowVisibilityEnsureFacts.FromResult(
                WorkbookWindowVisibilityEnsureResult.Create(
                    WorkbookWindowVisibilityEnsureOutcome.MadeVisible,
                    @"C:\cases\case.xlsx",
                    window: null,
                    elapsedMilliseconds: 0,
                    visibleAfterSet: false));
            TaskPaneRefreshAttemptResult attemptResult = CreateShownAttempt(null);

            VisibilityRecoveryOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(
                workbook: null,
                inputWindow: null,
                attemptResult: attemptResult,
                workbookWindowEnsureFacts: ensureFacts);

            Assert.Equal(VisibilityRecoveryOutcomeStatus.Degraded, outcome.Status);
            Assert.Equal("paneVisibleWithDegradedRecoveryFacts", outcome.Reason);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsPaneVisible);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal("workbookWindowEnsureVisibleAfterSet=False", outcome.DegradedReason);
        }

        [Fact]
        public void BuildRefreshSourceSelectionOutcome_WhenSnapshotNotReached_StaysNotReached()
        {
            RefreshSourceSelectionOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(
                TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied());

            Assert.Equal(RefreshSourceSelectionOutcomeStatus.NotReached, outcome.Status);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.CanContinueRefresh);
            Assert.False(outcome.IsCacheFallback);
            Assert.False(outcome.IsRebuildRequired);
        }

        [Fact]
        public void BuildRebuildFallbackOutcome_WhenSnapshotMissing_StaysSkippedContinuationOutcome()
        {
            RebuildFallbackOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildRebuildFallbackOutcome(
                TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied());

            Assert.Equal(RebuildFallbackOutcomeStatus.Skipped, outcome.Status);
            Assert.Equal("snapshotAcquisitionNotReached", outcome.Reason);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.CanContinueRefresh);
            Assert.False(outcome.IsRequired);
        }

        [Fact]
        public void FormatRefreshSourceSelectionAction_FreezesTraceActionNames()
        {
            Assert.Equal(
                "refresh-source-selected",
                TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(
                    TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(CreateShownAttempt(CreateBuildResult(
                        TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.CaseCache,
                        snapshotText: "snapshot")))));
            Assert.Equal(
                "refresh-source-fallback",
                TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(
                    TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(CreateShownAttempt(CreateBuildResult(
                        TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.BaseCacheFallback,
                        snapshotText: "snapshot",
                        fallbackReasons: "LatestMasterVersionUnavailable")))));
            Assert.Equal(
                "refresh-source-rebuild-required",
                TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(
                    TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(CreateShownAttempt(CreateBuildResult(
                        TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                        snapshotText: "snapshot",
                        fallbackReasons: "CaseCacheStale",
                        masterListRebuildAttempted: true,
                        masterListRebuildSucceeded: true)))));
            Assert.Equal(
                "refresh-source-failed",
                TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(
                    TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(TaskPaneRefreshAttemptResult.Failed(CreateBuildResult(
                        TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                        snapshotText: string.Empty,
                        fallbackReasons: "CacheUnavailable",
                        masterListRebuildAttempted: true,
                        failureReason: "SnapshotBuildException")))));
            Assert.Equal(
                "refresh-source-not-reached",
                TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(
                    TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied())));
            Assert.Equal(
                "refresh-source-unknown",
                TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(RefreshSourceSelectionOutcome.Unknown("notEvaluated")));
        }

        [Fact]
        public void FormatVisibilityRecoveryDetails_IncludesCompletionSourceAttemptAndDisplayCompletableFacts()
        {
            VisibilityRecoveryOutcome outcome = VisibilityRecoveryOutcome.Completed(
                "refreshedShown",
                VisibilityRecoveryTargetKind.ActiveWorkbookFallback,
                PaneVisibleSource.RefreshedShown,
                workbookWindowEnsureStatus: null,
                fullRecoveryAttempted: false,
                fullRecoverySucceeded: null);
            TaskPaneRefreshAttemptResult attemptResult = CreateShownAttempt(null);

            string details = TaskPaneNormalizedOutcomeMapper.FormatVisibilityRecoveryDetails(
                "KernelCasePresentationService.ShowCreatedCase.PostRelease",
                outcome,
                attemptResult,
                "ready-show-attempt",
                2,
                workbookWindowEnsureFacts: null);

            Assert.Contains("completionSource=ready-show-attempt", details);
            Assert.Contains("visibilityRecoveryStatus=Completed", details);
            Assert.Contains("visibilityRecoveryDisplayCompletable=True", details);
            Assert.Contains("visibilityRecoveryTargetKind=ActiveWorkbookFallback", details);
            Assert.Contains("attempt=2", details);
        }

        [Fact]
        public void FormatVisibilityRecoveryDetails_PreservesFieldOrder()
        {
            VisibilityRecoveryOutcome outcome = VisibilityRecoveryOutcome.Completed(
                "refreshedShown",
                VisibilityRecoveryTargetKind.ActiveWorkbookFallback,
                PaneVisibleSource.RefreshedShown,
                workbookWindowEnsureStatus: null,
                fullRecoveryAttempted: false,
                fullRecoverySucceeded: null);
            TaskPaneRefreshAttemptResult attemptResult = CreateShownAttempt(null);

            string details = TaskPaneNormalizedOutcomeMapper.FormatVisibilityRecoveryDetails(
                "KernelCasePresentationService.ShowCreatedCase.PostRelease",
                outcome,
                attemptResult,
                "ready-show-attempt",
                2,
                workbookWindowEnsureFacts: null);

            Assert.Equal(
                new[]
                {
                    "reason",
                    "completionSource",
                    "visibilityRecoveryStatus",
                    "visibilityRecoveryReason",
                    "visibilityRecoveryTerminal",
                    "visibilityRecoveryDisplayCompletable",
                    "visibilityRecoveryPaneVisible",
                    "visibilityRecoveryTargetKind",
                    "visibilityPaneVisibleSource",
                    "visibilityRecoveryDegradedReason",
                    "refreshSucceeded",
                    "refreshCompleted",
                    "preContextFullRecoveryAttempted",
                    "preContextFullRecoverySucceeded",
                    "attempt"
                },
                ExtractFieldNames(details));
            Assert.Contains("completionSource=ready-show-attempt", details);
            Assert.Contains("visibilityRecoveryStatus=Completed", details);
            Assert.Contains("visibilityRecoveryDisplayCompletable=True", details);
            Assert.Contains("attempt=2", details);
        }

        [Fact]
        public void FormatRefreshSourceSelectionDetails_PreservesFieldOrder()
        {
            TaskPaneRefreshAttemptResult attemptResult = CreateShownAttempt(CreateBuildResult(
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                snapshotText: "snapshot",
                fallbackReasons: "CaseCacheStale",
                masterListRebuildAttempted: true,
                masterListRebuildSucceeded: true));
            RefreshSourceSelectionOutcome outcome =
                TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(attemptResult);

            string details = TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionDetails(
                "KernelCasePresentationService.ShowCreatedCase.PostRelease",
                outcome,
                attemptResult,
                "refresh",
                1);

            Assert.Equal(
                new[]
                {
                    "reason",
                    "completionSource",
                    "refreshSourceStatus",
                    "selectedSource",
                    "selectionReason",
                    "fallbackReasons",
                    "refreshSourceTerminal",
                    "refreshSourceCanContinue",
                    "cacheFallback",
                    "rebuildRequired",
                    "masterListRebuildAttempted",
                    "masterListRebuildSucceeded",
                    "snapshotTextAvailable",
                    "updatedCaseSnapshotCache",
                    "failureReason",
                    "degradedReason",
                    "refreshSucceeded",
                    "refreshCompleted",
                    "paneVisible",
                    "attempt"
                },
                ExtractFieldNames(details));
            Assert.Contains("completionSource=refresh", details);
            Assert.Contains("refreshSourceStatus=RebuildRequired", details);
            Assert.Contains("rebuildRequired=True", details);
            Assert.Contains("attempt=1", details);
        }

        [Fact]
        public void FormatRebuildFallbackDetails_PreservesFieldOrder()
        {
            TaskPaneRefreshAttemptResult attemptResult = CreateShownAttempt(CreateBuildResult(
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild,
                snapshotText: "snapshot",
                fallbackReasons: "CaseCacheStale",
                masterListRebuildAttempted: true,
                masterListRebuildSucceeded: true));
            RebuildFallbackOutcome outcome =
                TaskPaneNormalizedOutcomeMapper.BuildRebuildFallbackOutcome(attemptResult);

            string details = TaskPaneNormalizedOutcomeMapper.FormatRebuildFallbackDetails(
                "KernelCasePresentationService.ShowCreatedCase.PostRelease",
                outcome,
                attemptResult,
                "refresh",
                1);

            Assert.Equal(
                new[]
                {
                    "reason",
                    "completionSource",
                    "rebuildFallbackStatus",
                    "rebuildFallbackRequired",
                    "rebuildFallbackTerminal",
                    "rebuildFallbackCanContinue",
                    "snapshotSource",
                    "fallbackReasons",
                    "masterListRebuildAttempted",
                    "masterListRebuildSucceeded",
                    "snapshotTextAvailable",
                    "updatedCaseSnapshotCache",
                    "failureReason",
                    "degradedReason",
                    "outcomeReason",
                    "refreshSucceeded",
                    "refreshCompleted",
                    "paneVisible",
                    "attempt"
                },
                ExtractFieldNames(details));
            Assert.Contains("completionSource=refresh", details);
            Assert.Contains("rebuildFallbackStatus=Completed", details);
            Assert.Contains("rebuildFallbackRequired=True", details);
            Assert.Contains("attempt=1", details);
        }

        private static TaskPaneRefreshAttemptResult CreateShownAttempt(
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult)
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

        private static TaskPaneSnapshotBuilderService.TaskPaneBuildResult CreateBuildResult(
            TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource snapshotSource,
            string snapshotText,
            string fallbackReasons = "",
            bool masterListRebuildAttempted = false,
            bool masterListRebuildSucceeded = false,
            string failureReason = "",
            string degradedReason = "")
        {
            return new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(
                snapshotText,
                updatedCaseSnapshotCache: false,
                snapshotSource,
                fallbackReasons,
                masterListRebuildAttempted,
                masterListRebuildSucceeded,
                failureReason,
                degradedReason);
        }

        private static string[] ExtractFieldNames(string details)
        {
            string[] fields = details.Split(',');
            string[] names = new string[fields.Length];
            for (int index = 0; index < fields.Length; index++)
            {
                int separator = fields[index].IndexOf('=');
                Assert.True(separator > 0, "Expected key=value field: " + fields[index]);
                names[index] = fields[index].Substring(0, separator);
            }

            return names;
        }
    }
}
