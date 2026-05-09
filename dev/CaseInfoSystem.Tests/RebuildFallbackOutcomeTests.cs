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
        public void ForegroundOutcome_WhenRequiredExecutionSucceeds_RemainsInputOnlyNotCompletionOwner()
        {
            ForegroundGuaranteeOutcome outcome = ForegroundGuaranteeOutcome.RequiredSucceeded(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                "foregroundRecoverySucceeded");
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(CreateDisplayCompletableVisibilityOutcome())
                .WithForegroundGuaranteeOutcome(outcome);

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, outcome.Status);
            Assert.True(outcome.WasRequired);
            Assert.True(outcome.WasExecutionAttempted);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal(ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow, outcome.TargetKind);
            Assert.True(outcome.RecoverySucceeded.GetValueOrDefault());
            Assert.True(SatisfiesCompletionHardGateInputContract(result));
        }

        [Fact]
        public void ForegroundOutcome_WhenRequiredExecutionReturnsFalse_DoesNotRoundToSuccessFailureOrCompletion()
        {
            ForegroundGuaranteeOutcome outcome = ForegroundGuaranteeOutcome.RequiredDegraded(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                "foregroundRecoveryReturnedFalse");
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(CreateDisplayCompletableVisibilityOutcome())
                .WithForegroundGuaranteeOutcome(outcome);

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredDegraded, outcome.Status);
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, outcome.Status);
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredFailed, outcome.Status);
            Assert.True(outcome.WasRequired);
            Assert.True(outcome.WasExecutionAttempted);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.False(outcome.RecoverySucceeded.GetValueOrDefault());
            Assert.True(SatisfiesCompletionHardGateInputContract(result));
        }

        [Fact]
        public void ForegroundOutcome_WhenRequiredFailed_BlocksCompletionHardGateInput()
        {
            ForegroundGuaranteeOutcome outcome = ForegroundGuaranteeOutcome.RequiredFailed(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                "foregroundRecoveryNotAttempted");
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(CreateDisplayCompletableVisibilityOutcome())
                .WithForegroundGuaranteeOutcome(outcome);

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredFailed, outcome.Status);
            Assert.True(outcome.WasRequired);
            Assert.False(outcome.WasExecutionAttempted);
            Assert.True(outcome.IsTerminal);
            Assert.False(outcome.IsDisplayCompletable);
            Assert.False(outcome.RecoverySucceeded.GetValueOrDefault());
            Assert.False(SatisfiesCompletionHardGateInputContract(result));
        }

        [Fact]
        public void ForegroundOutcome_WhenNotRequired_IsDisplayCompletableButNotForegroundSuccess()
        {
            ForegroundGuaranteeOutcome outcome = ForegroundGuaranteeOutcome.NotRequired("foregroundNotRequired");

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.NotRequired, outcome.Status);
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, outcome.Status);
            Assert.False(outcome.WasRequired);
            Assert.False(outcome.WasExecutionAttempted);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal(ForegroundGuaranteeTargetKind.NotRequired, outcome.TargetKind);
            Assert.False(outcome.RecoverySucceeded.HasValue);
        }

        [Fact]
        public void ForegroundOutcome_WhenSkippedAlreadyVisible_IsTerminalButNotForegroundSuccess()
        {
            ForegroundGuaranteeOutcome outcome = ForegroundGuaranteeOutcome.SkippedAlreadyVisible(
                "visibleCasePaneAlreadyShown");

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.SkippedAlreadyVisible, outcome.Status);
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, outcome.Status);
            Assert.False(outcome.WasRequired);
            Assert.False(outcome.WasExecutionAttempted);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.Equal(ForegroundGuaranteeTargetKind.AlreadyVisible, outcome.TargetKind);
            Assert.False(outcome.RecoverySucceeded.HasValue);
        }

        [Fact]
        public void ForegroundDisplayCompletableInputMapping_PreservesTerminalInputContract()
        {
            ForegroundGuaranteeOutcome[] displayCompletableInputs =
            {
                ForegroundGuaranteeOutcome.RequiredSucceeded(
                    ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                    "foregroundRecoverySucceeded"),
                ForegroundGuaranteeOutcome.RequiredDegraded(
                    ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                    "foregroundRecoveryReturnedFalse"),
                ForegroundGuaranteeOutcome.NotRequired("foregroundNotRequired"),
                ForegroundGuaranteeOutcome.SkippedAlreadyVisible("visibleCasePaneAlreadyShown"),
            };
            ForegroundGuaranteeOutcome[] nonDisplayCompletableInputs =
            {
                ForegroundGuaranteeOutcome.RequiredFailed(
                    ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                    "foregroundRecoveryNotAttempted"),
                ForegroundGuaranteeOutcome.SkippedNoKnownTarget("foregroundNoKnownTarget"),
                ForegroundGuaranteeOutcome.Unknown("pendingForegroundGuaranteeOutcome"),
            };

            foreach (ForegroundGuaranteeOutcome outcome in displayCompletableInputs)
            {
                Assert.True(outcome.IsTerminal);
                Assert.True(outcome.IsDisplayCompletable);
                Assert.True(IsForegroundDisplayCompletableTerminalInputContract(outcome));
            }

            foreach (ForegroundGuaranteeOutcome outcome in nonDisplayCompletableInputs)
            {
                Assert.False(outcome.IsDisplayCompletable);
                Assert.False(IsForegroundDisplayCompletableTerminalInputContract(outcome));
            }

            Assert.True(nonDisplayCompletableInputs[0].IsTerminal);
            Assert.True(nonDisplayCompletableInputs[1].IsTerminal);
            Assert.False(nonDisplayCompletableInputs[2].IsTerminal);
        }

        [Fact]
        public void ForegroundDisplayCompletableInput_RemainsInputOnlyNotSuccessFailureOrDirectCompletion()
        {
            ForegroundGuaranteeOutcome degradedOutcome = ForegroundGuaranteeOutcome.RequiredDegraded(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                "foregroundRecoveryReturnedFalse");
            ForegroundGuaranteeOutcome notRequiredOutcome = ForegroundGuaranteeOutcome.NotRequired("foregroundNotRequired");
            ForegroundGuaranteeOutcome skippedAlreadyVisibleOutcome = ForegroundGuaranteeOutcome.SkippedAlreadyVisible(
                "visibleCasePaneAlreadyShown");
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(CreateDisplayCompletableVisibilityOutcome())
                .WithForegroundGuaranteeOutcome(degradedOutcome);

            Assert.True(IsForegroundDisplayCompletableTerminalInputContract(degradedOutcome));
            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredDegraded, degradedOutcome.Status);
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, degradedOutcome.Status);
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredFailed, degradedOutcome.Status);
            Assert.False(degradedOutcome.RecoverySucceeded.GetValueOrDefault());
            Assert.True(SatisfiesCompletionHardGateInputContract(result));

            Assert.True(IsForegroundDisplayCompletableTerminalInputContract(notRequiredOutcome));
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, notRequiredOutcome.Status);
            Assert.False(notRequiredOutcome.WasRequired);
            Assert.False(notRequiredOutcome.WasExecutionAttempted);
            Assert.False(notRequiredOutcome.RecoverySucceeded.HasValue);

            Assert.True(IsForegroundDisplayCompletableTerminalInputContract(skippedAlreadyVisibleOutcome));
            Assert.NotEqual(ForegroundGuaranteeOutcomeStatus.RequiredSucceeded, skippedAlreadyVisibleOutcome.Status);
            Assert.False(skippedAlreadyVisibleOutcome.WasRequired);
            Assert.False(skippedAlreadyVisibleOutcome.WasExecutionAttempted);
            Assert.False(skippedAlreadyVisibleOutcome.RecoverySucceeded.HasValue);
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
            Assert.False(SatisfiesCompletionHardGateInputContract(result));
        }

        [Fact]
        public void CompletionGateInput_WhenVisibilityIsNotDisplayCompletable_RemainsBlocked()
        {
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome.Failed(
                    "paneVisible=false",
                    VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                    PaneVisibleSource.None,
                    workbookWindowEnsureStatus: null,
                    fullRecoveryAttempted: false,
                    fullRecoverySucceeded: null))
                .WithForegroundGuaranteeOutcome(ForegroundGuaranteeOutcome.NotRequired("foregroundNotRequired"));

            Assert.True(result.IsRefreshSucceeded);
            Assert.True(result.IsPaneVisible);
            Assert.True(result.VisibilityRecoveryOutcome.IsTerminal);
            Assert.False(result.VisibilityRecoveryOutcome.IsDisplayCompletable);
            Assert.True(result.IsForegroundGuaranteeTerminal);
            Assert.True(result.ForegroundGuaranteeOutcome.IsDisplayCompletable);
            Assert.False(SatisfiesCompletionHardGateInputContract(result));
        }

        [Fact]
        public void CompletionGateInput_WhenForegroundIsNotTerminal_RemainsBlocked()
        {
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(CreateDisplayCompletableVisibilityOutcome())
                .WithForegroundGuaranteeOutcome(ForegroundGuaranteeOutcome.Unknown("pendingForegroundGuaranteeOutcome"));

            Assert.True(result.IsRefreshSucceeded);
            Assert.True(result.IsPaneVisible);
            Assert.True(result.VisibilityRecoveryOutcome.IsDisplayCompletable);
            Assert.False(result.IsForegroundGuaranteeTerminal);
            Assert.False(result.ForegroundGuaranteeOutcome.IsDisplayCompletable);
            Assert.False(SatisfiesCompletionHardGateInputContract(result));
        }

        [Fact]
        public void CompletionGateInput_WhenAllRequiredFactsAreDisplayCompletable_IsEligibleOnlyAsInput()
        {
            TaskPaneRefreshAttemptResult result = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(CreateDisplayCompletableVisibilityOutcome())
                .WithForegroundGuaranteeOutcome(ForegroundGuaranteeOutcome.RequiredDegraded(
                    ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                    "foregroundRecoveryReturnedFalse"));

            Assert.True(result.IsRefreshSucceeded);
            Assert.True(result.IsPaneVisible);
            Assert.True(result.VisibilityRecoveryOutcome.IsTerminal);
            Assert.True(result.VisibilityRecoveryOutcome.IsDisplayCompletable);
            Assert.True(result.IsForegroundGuaranteeTerminal);
            Assert.True(result.ForegroundGuaranteeOutcome.IsDisplayCompletable);
            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredDegraded, result.ForegroundGuaranteeOutcome.Status);
            Assert.True(SatisfiesCompletionHardGateInputContract(result));
        }

        private static VisibilityRecoveryOutcome CreateDisplayCompletableVisibilityOutcome()
        {
            return VisibilityRecoveryOutcome.Completed(
                "refreshedShown",
                VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                PaneVisibleSource.RefreshedShown,
                WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible,
                fullRecoveryAttempted: false,
                fullRecoverySucceeded: null);
        }

        private static bool IsForegroundDisplayCompletableTerminalInputContract(ForegroundGuaranteeOutcome outcome)
        {
            return outcome != null
                && outcome.IsTerminal
                && outcome.IsDisplayCompletable;
        }

        private static bool SatisfiesCompletionHardGateInputContract(TaskPaneRefreshAttemptResult result)
        {
            return result != null
                && result.IsRefreshSucceeded
                && result.IsPaneVisible
                && result.VisibilityRecoveryOutcome != null
                && result.VisibilityRecoveryOutcome.IsTerminal
                && result.VisibilityRecoveryOutcome.IsDisplayCompletable
                && result.IsForegroundGuaranteeTerminal
                && result.ForegroundGuaranteeOutcome != null
                && result.ForegroundGuaranteeOutcome.IsDisplayCompletable;
        }
    }
}
