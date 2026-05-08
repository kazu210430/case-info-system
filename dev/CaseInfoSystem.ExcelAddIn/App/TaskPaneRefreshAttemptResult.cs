using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum ForegroundGuaranteeOutcomeStatus
    {
        Unknown = 0,
        NotRequired = 1,
        SkippedAlreadyVisible = 2,
        SkippedNoKnownTarget = 3,
        RequiredSucceeded = 4,
        RequiredDegraded = 5,
        RequiredFailed = 6,
    }

    internal enum ForegroundGuaranteeTargetKind
    {
        Unknown = 0,
        NotRequired = 1,
        AlreadyVisible = 2,
        ExplicitWorkbookWindow = 3,
        ActiveWorkbookFallback = 4,
        NoKnownTarget = 5,
    }

    internal enum VisibilityRecoveryOutcomeStatus
    {
        Unknown = 0,
        Completed = 1,
        Skipped = 2,
        Degraded = 3,
        Failed = 4,
    }

    internal enum VisibilityRecoveryTargetKind
    {
        Unknown = 0,
        ExplicitWorkbookWindow = 1,
        ActiveWorkbookFallback = 2,
        AlreadyVisible = 3,
        NoKnownTarget = 4,
    }

    internal enum PaneVisibleSource
    {
        None = 0,
        AlreadyVisibleHost = 1,
        ReusedShown = 2,
        RefreshedShown = 3,
        Unknown = 4,
    }

    internal enum RebuildFallbackOutcomeStatus
    {
        Unknown = 0,
        Skipped = 1,
        Completed = 2,
        Degraded = 3,
        Failed = 4,
    }

    internal enum RefreshSourceSelectionOutcomeStatus
    {
        Unknown = 0,
        NotReached = 1,
        Selected = 2,
        DegradedSelected = 3,
        FallbackSelected = 4,
        RebuildRequired = 5,
        Failed = 6,
    }

    internal sealed class RefreshSourceSelectionOutcome
    {
        private RefreshSourceSelectionOutcome(
            RefreshSourceSelectionOutcomeStatus status,
            TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource selectedSource,
            string selectionReason,
            string fallbackReasons,
            bool isTerminal,
            bool canContinueRefresh,
            bool isCacheFallback,
            bool isRebuildRequired,
            bool masterListRebuildAttempted,
            bool masterListRebuildSucceeded,
            bool snapshotTextAvailable,
            bool updatedCaseSnapshotCache,
            string failureReason,
            string degradedReason)
        {
            Status = status;
            SelectedSource = selectedSource;
            SelectionReason = selectionReason ?? string.Empty;
            FallbackReasons = fallbackReasons ?? string.Empty;
            IsTerminal = isTerminal;
            CanContinueRefresh = canContinueRefresh;
            IsCacheFallback = isCacheFallback;
            IsRebuildRequired = isRebuildRequired;
            MasterListRebuildAttempted = masterListRebuildAttempted;
            MasterListRebuildSucceeded = masterListRebuildSucceeded;
            SnapshotTextAvailable = snapshotTextAvailable;
            UpdatedCaseSnapshotCache = updatedCaseSnapshotCache;
            FailureReason = failureReason ?? string.Empty;
            DegradedReason = degradedReason ?? string.Empty;
        }

        internal RefreshSourceSelectionOutcomeStatus Status { get; }

        internal TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource SelectedSource { get; }

        internal string SelectionReason { get; }

        internal string FallbackReasons { get; }

        internal bool IsTerminal { get; }

        internal bool CanContinueRefresh { get; }

        internal bool IsCacheFallback { get; }

        internal bool IsRebuildRequired { get; }

        internal bool MasterListRebuildAttempted { get; }

        internal bool MasterListRebuildSucceeded { get; }

        internal bool SnapshotTextAvailable { get; }

        internal bool UpdatedCaseSnapshotCache { get; }

        internal string FailureReason { get; }

        internal string DegradedReason { get; }

        internal static RefreshSourceSelectionOutcome Unknown(string reason)
        {
            return new RefreshSourceSelectionOutcome(
                RefreshSourceSelectionOutcomeStatus.Unknown,
                TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None,
                reason,
                string.Empty,
                isTerminal: false,
                canContinueRefresh: false,
                isCacheFallback: false,
                isRebuildRequired: false,
                masterListRebuildAttempted: false,
                masterListRebuildSucceeded: false,
                snapshotTextAvailable: false,
                updatedCaseSnapshotCache: false,
                failureReason: string.Empty,
                degradedReason: string.Empty);
        }

        internal static RefreshSourceSelectionOutcome FromAttemptResult(TaskPaneRefreshAttemptResult attemptResult)
        {
            if (attemptResult == null)
            {
                return Unknown("attemptResultMissing");
            }

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult = attemptResult.SnapshotBuildResult;
            if (buildResult == null)
            {
                return new RefreshSourceSelectionOutcome(
                    RefreshSourceSelectionOutcomeStatus.NotReached,
                    TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None,
                    "snapshotAcquisitionNotReached",
                    string.Empty,
                    isTerminal: true,
                    canContinueRefresh: attemptResult.IsRefreshSucceeded,
                    isCacheFallback: false,
                    isRebuildRequired: false,
                    masterListRebuildAttempted: false,
                    masterListRebuildSucceeded: false,
                    snapshotTextAvailable: false,
                    updatedCaseSnapshotCache: false,
                    failureReason: string.Empty,
                    degradedReason: string.Empty);
            }

            TaskPaneSnapshotBuilderService.SnapshotSourceSelectionFacts facts = buildResult.SourceSelectionFacts;
            if (facts == null)
            {
                return Unknown("sourceSelectionFactsMissing");
            }

            if (!facts.SnapshotTextAvailable)
            {
                return new RefreshSourceSelectionOutcome(
                    RefreshSourceSelectionOutcomeStatus.Failed,
                    facts.SelectedSource,
                    string.IsNullOrWhiteSpace(facts.SelectionReason) ? "NoSnapshotText" : facts.SelectionReason,
                    facts.FallbackReasons,
                    isTerminal: true,
                    canContinueRefresh: false,
                    isCacheFallback: facts.IsCacheFallback,
                    isRebuildRequired: facts.IsRebuildRequired,
                    masterListRebuildAttempted: facts.MasterListRebuildAttempted,
                    masterListRebuildSucceeded: facts.MasterListRebuildSucceeded,
                    snapshotTextAvailable: false,
                    updatedCaseSnapshotCache: facts.UpdatedCaseSnapshotCache,
                    failureReason: string.IsNullOrWhiteSpace(facts.FailureReason) ? "NoSnapshotText" : facts.FailureReason,
                    degradedReason: facts.DegradedReason);
            }

            RefreshSourceSelectionOutcomeStatus status = ResolveStatus(facts);
            return new RefreshSourceSelectionOutcome(
                status,
                facts.SelectedSource,
                facts.SelectionReason,
                facts.FallbackReasons,
                isTerminal: true,
                canContinueRefresh: true,
                isCacheFallback: facts.IsCacheFallback,
                isRebuildRequired: facts.IsRebuildRequired,
                masterListRebuildAttempted: facts.MasterListRebuildAttempted,
                masterListRebuildSucceeded: facts.MasterListRebuildSucceeded,
                snapshotTextAvailable: true,
                updatedCaseSnapshotCache: facts.UpdatedCaseSnapshotCache,
                failureReason: facts.FailureReason,
                degradedReason: facts.DegradedReason);
        }

        private static RefreshSourceSelectionOutcomeStatus ResolveStatus(TaskPaneSnapshotBuilderService.SnapshotSourceSelectionFacts facts)
        {
            if (!string.IsNullOrWhiteSpace(facts.DegradedReason)
                || (!string.IsNullOrWhiteSpace(facts.FailureReason) && facts.SnapshotTextAvailable))
            {
                return RefreshSourceSelectionOutcomeStatus.DegradedSelected;
            }

            if (facts.IsCacheFallback)
            {
                return RefreshSourceSelectionOutcomeStatus.FallbackSelected;
            }

            if (facts.IsRebuildRequired)
            {
                return RefreshSourceSelectionOutcomeStatus.RebuildRequired;
            }

            return RefreshSourceSelectionOutcomeStatus.Selected;
        }
    }

    internal sealed class RebuildFallbackOutcome
    {
        private RebuildFallbackOutcome(
            RebuildFallbackOutcomeStatus status,
            bool isRequired,
            bool isTerminal,
            bool canContinueRefresh,
            TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource snapshotSource,
            string fallbackReasons,
            bool masterListRebuildAttempted,
            bool masterListRebuildSucceeded,
            bool snapshotTextAvailable,
            bool updatedCaseSnapshotCache,
            string failureReason,
            string degradedReason,
            string reason)
        {
            Status = status;
            IsRequired = isRequired;
            IsTerminal = isTerminal;
            CanContinueRefresh = canContinueRefresh;
            SnapshotSource = snapshotSource;
            FallbackReasons = fallbackReasons ?? string.Empty;
            MasterListRebuildAttempted = masterListRebuildAttempted;
            MasterListRebuildSucceeded = masterListRebuildSucceeded;
            SnapshotTextAvailable = snapshotTextAvailable;
            UpdatedCaseSnapshotCache = updatedCaseSnapshotCache;
            FailureReason = failureReason ?? string.Empty;
            DegradedReason = degradedReason ?? string.Empty;
            Reason = reason ?? string.Empty;
        }

        internal RebuildFallbackOutcomeStatus Status { get; }

        internal bool IsRequired { get; }

        internal bool IsTerminal { get; }

        internal bool CanContinueRefresh { get; }

        internal TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource SnapshotSource { get; }

        internal string FallbackReasons { get; }

        internal bool MasterListRebuildAttempted { get; }

        internal bool MasterListRebuildSucceeded { get; }

        internal bool SnapshotTextAvailable { get; }

        internal bool UpdatedCaseSnapshotCache { get; }

        internal string FailureReason { get; }

        internal string DegradedReason { get; }

        internal string Reason { get; }

        internal static RebuildFallbackOutcome Unknown(string reason)
        {
            return new RebuildFallbackOutcome(
                RebuildFallbackOutcomeStatus.Unknown,
                isRequired: false,
                isTerminal: false,
                canContinueRefresh: false,
                snapshotSource: TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None,
                fallbackReasons: string.Empty,
                masterListRebuildAttempted: false,
                masterListRebuildSucceeded: false,
                snapshotTextAvailable: false,
                updatedCaseSnapshotCache: false,
                failureReason: string.Empty,
                degradedReason: string.Empty,
                reason: reason);
        }

        internal static RebuildFallbackOutcome FromBuildResult(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult)
        {
            if (buildResult == null)
            {
                return Skipped(
                    "snapshotAcquisitionNotReached",
                    TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None,
                    string.Empty,
                    updatedCaseSnapshotCache: false);
            }

            if (buildResult.SnapshotSource == TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild
                || buildResult.MasterListRebuildAttempted)
            {
                if (buildResult.MasterListRebuildSucceeded && buildResult.SnapshotTextAvailable)
                {
                    return Completed(buildResult, "masterListRebuildCompleted");
                }

                if (buildResult.SnapshotTextAvailable)
                {
                    return Degraded(
                        buildResult,
                        string.IsNullOrWhiteSpace(buildResult.DegradedReason)
                            ? "masterListRebuildReturnedFallbackSnapshot"
                            : buildResult.DegradedReason);
                }

                return Failed(
                    buildResult,
                    string.IsNullOrWhiteSpace(buildResult.FailureReason)
                        ? "NoSnapshotText"
                        : buildResult.FailureReason);
            }

            if (buildResult.SnapshotTextAvailable)
            {
                return Skipped(
                    "snapshotSource=" + buildResult.SnapshotSource.ToString(),
                    buildResult.SnapshotSource,
                    buildResult.FallbackReasons,
                    buildResult.UpdatedCaseSnapshotCache);
            }

            if (!string.IsNullOrWhiteSpace(buildResult.FailureReason))
            {
                return Failed(buildResult, buildResult.FailureReason);
            }

            return Skipped(
                "snapshotAcquisitionSkipped",
                buildResult.SnapshotSource,
                buildResult.FallbackReasons,
                buildResult.UpdatedCaseSnapshotCache);
        }

        private static RebuildFallbackOutcome Skipped(
            string reason,
            TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource snapshotSource,
            string fallbackReasons,
            bool updatedCaseSnapshotCache)
        {
            return new RebuildFallbackOutcome(
                RebuildFallbackOutcomeStatus.Skipped,
                isRequired: false,
                isTerminal: true,
                canContinueRefresh: true,
                snapshotSource: snapshotSource,
                fallbackReasons: fallbackReasons,
                masterListRebuildAttempted: false,
                masterListRebuildSucceeded: false,
                snapshotTextAvailable: snapshotSource != TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None,
                updatedCaseSnapshotCache: updatedCaseSnapshotCache,
                failureReason: string.Empty,
                degradedReason: string.Empty,
                reason: reason);
        }

        private static RebuildFallbackOutcome Completed(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult, string reason)
        {
            return new RebuildFallbackOutcome(
                RebuildFallbackOutcomeStatus.Completed,
                isRequired: true,
                isTerminal: true,
                canContinueRefresh: true,
                snapshotSource: buildResult.SnapshotSource,
                fallbackReasons: buildResult.FallbackReasons,
                masterListRebuildAttempted: buildResult.MasterListRebuildAttempted,
                masterListRebuildSucceeded: buildResult.MasterListRebuildSucceeded,
                snapshotTextAvailable: buildResult.SnapshotTextAvailable,
                updatedCaseSnapshotCache: buildResult.UpdatedCaseSnapshotCache,
                failureReason: buildResult.FailureReason,
                degradedReason: buildResult.DegradedReason,
                reason: reason);
        }

        private static RebuildFallbackOutcome Degraded(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult, string degradedReason)
        {
            return new RebuildFallbackOutcome(
                RebuildFallbackOutcomeStatus.Degraded,
                isRequired: true,
                isTerminal: true,
                canContinueRefresh: true,
                snapshotSource: buildResult.SnapshotSource,
                fallbackReasons: buildResult.FallbackReasons,
                masterListRebuildAttempted: buildResult.MasterListRebuildAttempted,
                masterListRebuildSucceeded: buildResult.MasterListRebuildSucceeded,
                snapshotTextAvailable: buildResult.SnapshotTextAvailable,
                updatedCaseSnapshotCache: buildResult.UpdatedCaseSnapshotCache,
                failureReason: buildResult.FailureReason,
                degradedReason: degradedReason,
                reason: "masterListRebuildDegraded");
        }

        private static RebuildFallbackOutcome Failed(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult, string failureReason)
        {
            return new RebuildFallbackOutcome(
                RebuildFallbackOutcomeStatus.Failed,
                isRequired: buildResult != null && (buildResult.SnapshotSource == TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild || buildResult.MasterListRebuildAttempted),
                isTerminal: true,
                canContinueRefresh: false,
                snapshotSource: buildResult == null ? TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None : buildResult.SnapshotSource,
                fallbackReasons: buildResult == null ? string.Empty : buildResult.FallbackReasons,
                masterListRebuildAttempted: buildResult != null && buildResult.MasterListRebuildAttempted,
                masterListRebuildSucceeded: buildResult != null && buildResult.MasterListRebuildSucceeded,
                snapshotTextAvailable: buildResult != null && buildResult.SnapshotTextAvailable,
                updatedCaseSnapshotCache: buildResult != null && buildResult.UpdatedCaseSnapshotCache,
                failureReason: failureReason,
                degradedReason: string.Empty,
                reason: "masterListRebuildFailed");
        }
    }

    internal sealed class VisibilityRecoveryOutcome
    {
        private VisibilityRecoveryOutcome(
            VisibilityRecoveryOutcomeStatus status,
            string reason,
            bool isTerminal,
            bool isPaneVisible,
            bool isDisplayCompletable,
            VisibilityRecoveryTargetKind targetKind,
            PaneVisibleSource paneVisibleSource,
            WorkbookWindowVisibilityEnsureOutcome? workbookWindowEnsureStatus,
            bool fullRecoveryAttempted,
            bool? fullRecoverySucceeded,
            string degradedReason)
        {
            Status = status;
            Reason = reason ?? string.Empty;
            IsTerminal = isTerminal;
            IsPaneVisible = isPaneVisible;
            IsDisplayCompletable = isDisplayCompletable;
            TargetKind = targetKind;
            PaneVisibleSource = paneVisibleSource;
            WorkbookWindowEnsureStatus = workbookWindowEnsureStatus;
            FullRecoveryAttempted = fullRecoveryAttempted;
            FullRecoverySucceeded = fullRecoverySucceeded;
            DegradedReason = degradedReason ?? string.Empty;
        }

        internal VisibilityRecoveryOutcomeStatus Status { get; }

        internal string Reason { get; }

        internal bool IsTerminal { get; }

        internal bool IsPaneVisible { get; }

        internal bool IsDisplayCompletable { get; }

        internal VisibilityRecoveryTargetKind TargetKind { get; }

        internal PaneVisibleSource PaneVisibleSource { get; }

        internal WorkbookWindowVisibilityEnsureOutcome? WorkbookWindowEnsureStatus { get; }

        internal bool FullRecoveryAttempted { get; }

        internal bool? FullRecoverySucceeded { get; }

        internal string DegradedReason { get; }

        internal static VisibilityRecoveryOutcome Unknown(string reason)
        {
            return new VisibilityRecoveryOutcome(
                VisibilityRecoveryOutcomeStatus.Unknown,
                reason,
                isTerminal: false,
                isPaneVisible: false,
                isDisplayCompletable: false,
                targetKind: VisibilityRecoveryTargetKind.Unknown,
                paneVisibleSource: PaneVisibleSource.Unknown,
                workbookWindowEnsureStatus: null,
                fullRecoveryAttempted: false,
                fullRecoverySucceeded: null,
                degradedReason: string.Empty);
        }

        internal static VisibilityRecoveryOutcome Completed(
            string reason,
            VisibilityRecoveryTargetKind targetKind,
            PaneVisibleSource paneVisibleSource,
            WorkbookWindowVisibilityEnsureOutcome? workbookWindowEnsureStatus,
            bool fullRecoveryAttempted,
            bool? fullRecoverySucceeded)
        {
            return new VisibilityRecoveryOutcome(
                VisibilityRecoveryOutcomeStatus.Completed,
                reason,
                isTerminal: true,
                isPaneVisible: true,
                isDisplayCompletable: true,
                targetKind: targetKind,
                paneVisibleSource: paneVisibleSource,
                workbookWindowEnsureStatus: workbookWindowEnsureStatus,
                fullRecoveryAttempted: fullRecoveryAttempted,
                fullRecoverySucceeded: fullRecoverySucceeded,
                degradedReason: string.Empty);
        }

        internal static VisibilityRecoveryOutcome Skipped(
            string reason,
            bool isPaneVisible,
            bool isDisplayCompletable,
            VisibilityRecoveryTargetKind targetKind,
            PaneVisibleSource paneVisibleSource,
            WorkbookWindowVisibilityEnsureOutcome? workbookWindowEnsureStatus,
            bool fullRecoveryAttempted,
            bool? fullRecoverySucceeded)
        {
            return new VisibilityRecoveryOutcome(
                VisibilityRecoveryOutcomeStatus.Skipped,
                reason,
                isTerminal: true,
                isPaneVisible: isPaneVisible,
                isDisplayCompletable: isDisplayCompletable,
                targetKind: targetKind,
                paneVisibleSource: paneVisibleSource,
                workbookWindowEnsureStatus: workbookWindowEnsureStatus,
                fullRecoveryAttempted: fullRecoveryAttempted,
                fullRecoverySucceeded: fullRecoverySucceeded,
                degradedReason: string.Empty);
        }

        internal static VisibilityRecoveryOutcome Degraded(
            string reason,
            VisibilityRecoveryTargetKind targetKind,
            PaneVisibleSource paneVisibleSource,
            WorkbookWindowVisibilityEnsureOutcome? workbookWindowEnsureStatus,
            bool fullRecoveryAttempted,
            bool? fullRecoverySucceeded,
            string degradedReason)
        {
            return new VisibilityRecoveryOutcome(
                VisibilityRecoveryOutcomeStatus.Degraded,
                reason,
                isTerminal: true,
                isPaneVisible: true,
                isDisplayCompletable: true,
                targetKind: targetKind,
                paneVisibleSource: paneVisibleSource,
                workbookWindowEnsureStatus: workbookWindowEnsureStatus,
                fullRecoveryAttempted: fullRecoveryAttempted,
                fullRecoverySucceeded: fullRecoverySucceeded,
                degradedReason: degradedReason);
        }

        internal static VisibilityRecoveryOutcome Failed(
            string reason,
            VisibilityRecoveryTargetKind targetKind,
            PaneVisibleSource paneVisibleSource,
            WorkbookWindowVisibilityEnsureOutcome? workbookWindowEnsureStatus,
            bool fullRecoveryAttempted,
            bool? fullRecoverySucceeded)
        {
            return new VisibilityRecoveryOutcome(
                VisibilityRecoveryOutcomeStatus.Failed,
                reason,
                isTerminal: true,
                isPaneVisible: false,
                isDisplayCompletable: false,
                targetKind: targetKind,
                paneVisibleSource: paneVisibleSource,
                workbookWindowEnsureStatus: workbookWindowEnsureStatus,
                fullRecoveryAttempted: fullRecoveryAttempted,
                fullRecoverySucceeded: fullRecoverySucceeded,
                degradedReason: string.Empty);
        }
    }

    internal sealed class ForegroundGuaranteeOutcome
    {
        private ForegroundGuaranteeOutcome(
            ForegroundGuaranteeOutcomeStatus status,
            bool wasRequired,
            bool wasExecutionAttempted,
            bool isTerminal,
            bool isDisplayCompletable,
            ForegroundGuaranteeTargetKind targetKind,
            bool? recoverySucceeded,
            string reason)
        {
            Status = status;
            WasRequired = wasRequired;
            WasExecutionAttempted = wasExecutionAttempted;
            IsTerminal = isTerminal;
            IsDisplayCompletable = isDisplayCompletable;
            TargetKind = targetKind;
            RecoverySucceeded = recoverySucceeded;
            Reason = reason ?? string.Empty;
        }

        internal ForegroundGuaranteeOutcomeStatus Status { get; }

        internal bool WasRequired { get; }

        internal bool WasExecutionAttempted { get; }

        internal bool IsTerminal { get; }

        internal bool IsDisplayCompletable { get; }

        internal ForegroundGuaranteeTargetKind TargetKind { get; }

        internal bool? RecoverySucceeded { get; }

        internal string Reason { get; }

        internal static ForegroundGuaranteeOutcome Unknown(string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.Unknown,
                wasRequired: false,
                wasExecutionAttempted: false,
                isTerminal: false,
                isDisplayCompletable: false,
                targetKind: ForegroundGuaranteeTargetKind.Unknown,
                recoverySucceeded: null,
                reason: reason);
        }

        internal static ForegroundGuaranteeOutcome NotRequired(string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.NotRequired,
                wasRequired: false,
                wasExecutionAttempted: false,
                isTerminal: true,
                isDisplayCompletable: true,
                targetKind: ForegroundGuaranteeTargetKind.NotRequired,
                recoverySucceeded: null,
                reason: reason);
        }

        internal static ForegroundGuaranteeOutcome SkippedAlreadyVisible(string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.SkippedAlreadyVisible,
                wasRequired: false,
                wasExecutionAttempted: false,
                isTerminal: true,
                isDisplayCompletable: true,
                targetKind: ForegroundGuaranteeTargetKind.AlreadyVisible,
                recoverySucceeded: null,
                reason: reason);
        }

        internal static ForegroundGuaranteeOutcome SkippedNoKnownTarget(string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.SkippedNoKnownTarget,
                wasRequired: false,
                wasExecutionAttempted: false,
                isTerminal: true,
                isDisplayCompletable: false,
                targetKind: ForegroundGuaranteeTargetKind.NoKnownTarget,
                recoverySucceeded: null,
                reason: reason);
        }

        internal static ForegroundGuaranteeOutcome RequiredSucceeded(ForegroundGuaranteeTargetKind targetKind, string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.RequiredSucceeded,
                wasRequired: true,
                wasExecutionAttempted: true,
                isTerminal: true,
                isDisplayCompletable: true,
                targetKind: targetKind,
                recoverySucceeded: true,
                reason: reason);
        }

        internal static ForegroundGuaranteeOutcome RequiredDegraded(ForegroundGuaranteeTargetKind targetKind, string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.RequiredDegraded,
                wasRequired: true,
                wasExecutionAttempted: true,
                isTerminal: true,
                isDisplayCompletable: true,
                targetKind: targetKind,
                recoverySucceeded: false,
                reason: reason);
        }

        internal static ForegroundGuaranteeOutcome RequiredFailed(ForegroundGuaranteeTargetKind targetKind, string reason)
        {
            return new ForegroundGuaranteeOutcome(
                ForegroundGuaranteeOutcomeStatus.RequiredFailed,
                wasRequired: true,
                wasExecutionAttempted: false,
                isTerminal: true,
                isDisplayCompletable: false,
                targetKind: targetKind,
                recoverySucceeded: false,
                reason: reason);
        }
    }

    internal sealed class ForegroundGuaranteeExecutionResult
    {
        internal ForegroundGuaranteeExecutionResult(bool executionAttempted, bool recovered, long elapsedMilliseconds)
        {
            ExecutionAttempted = executionAttempted;
            Recovered = recovered;
            ElapsedMilliseconds = elapsedMilliseconds;
        }

        internal bool ExecutionAttempted { get; }

        internal bool Recovered { get; }

        internal long ElapsedMilliseconds { get; }

        internal static ForegroundGuaranteeExecutionResult NotAttempted()
        {
            return new ForegroundGuaranteeExecutionResult(false, recovered: false, elapsedMilliseconds: 0);
        }
    }

    internal sealed class TaskPaneRefreshAttemptResult
    {
        private TaskPaneRefreshAttemptResult(
            bool isRefreshSucceeded,
            bool wasSkipped = false,
            bool wasContextRejected = false,
            bool isPaneVisible = false,
            bool isRefreshCompleted = false,
            string completionBasis = "",
            ForegroundGuaranteeOutcome foregroundGuaranteeOutcome = null,
            WorkbookContext foregroundContext = null,
            Excel.Workbook foregroundWorkbook = null,
            Excel.Window foregroundWindow = null,
            bool isForegroundRecoveryServiceAvailable = false,
            VisibilityRecoveryOutcome visibilityRecoveryOutcome = null,
            RebuildFallbackOutcome rebuildFallbackOutcome = null,
            RefreshSourceSelectionOutcome refreshSourceSelectionOutcome = null,
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult snapshotBuildResult = null,
            PaneVisibleSource paneVisibleSource = PaneVisibleSource.None,
            bool preContextRecoveryAttempted = false,
            bool? preContextRecoverySucceeded = null)
        {
            IsRefreshSucceeded = isRefreshSucceeded;
            WasSkipped = wasSkipped;
            WasContextRejected = wasContextRejected;
            IsPaneVisible = isPaneVisible;
            IsRefreshCompleted = isRefreshCompleted;
            CompletionBasis = completionBasis ?? string.Empty;
            ForegroundGuaranteeOutcome = foregroundGuaranteeOutcome ?? ForegroundGuaranteeOutcome.Unknown("notEvaluated");
            ForegroundContext = foregroundContext;
            ForegroundWorkbook = foregroundWorkbook;
            ForegroundWindow = foregroundWindow;
            IsForegroundRecoveryServiceAvailable = isForegroundRecoveryServiceAvailable;
            VisibilityRecoveryOutcome = visibilityRecoveryOutcome ?? VisibilityRecoveryOutcome.Unknown("notEvaluated");
            RebuildFallbackOutcome = rebuildFallbackOutcome ?? RebuildFallbackOutcome.Unknown("notEvaluated");
            RefreshSourceSelectionOutcome = refreshSourceSelectionOutcome ?? RefreshSourceSelectionOutcome.Unknown("notEvaluated");
            SnapshotBuildResult = snapshotBuildResult;
            PaneVisibleSource = paneVisibleSource;
            PreContextRecoveryAttempted = preContextRecoveryAttempted;
            PreContextRecoverySucceeded = preContextRecoverySucceeded;
        }

        internal static TaskPaneRefreshAttemptResult Succeeded()
        {
            return Succeeded(
                foregroundGuaranteeRequired: false,
                completionBasis: "refreshCompleted");
        }

        internal static TaskPaneRefreshAttemptResult Succeeded(bool foregroundGuaranteeRequired, string completionBasis)
        {
            return new TaskPaneRefreshAttemptResult(
                true,
                isPaneVisible: true,
                isRefreshCompleted: true,
                completionBasis: completionBasis,
                foregroundGuaranteeOutcome: foregroundGuaranteeRequired
                    ? ForegroundGuaranteeOutcome.RequiredSucceeded(
                        ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                        completionBasis)
                    : ForegroundGuaranteeOutcome.NotRequired(completionBasis),
                paneVisibleSource: PaneVisibleSource.RefreshedShown);
        }

        internal static TaskPaneRefreshAttemptResult RefreshCompletedPendingForeground(
            WorkbookContext foregroundContext,
            Excel.Workbook foregroundWorkbook,
            Excel.Window foregroundWindow,
            bool isForegroundRecoveryServiceAvailable,
            string completionBasis,
            PaneVisibleSource paneVisibleSource,
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult snapshotBuildResult,
            bool preContextRecoveryAttempted,
            bool? preContextRecoverySucceeded)
        {
            return new TaskPaneRefreshAttemptResult(
                true,
                isPaneVisible: true,
                isRefreshCompleted: true,
                completionBasis: completionBasis,
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.Unknown("pendingForegroundGuaranteeOutcome"),
                foregroundContext: foregroundContext,
                foregroundWorkbook: foregroundWorkbook,
                foregroundWindow: foregroundWindow,
                isForegroundRecoveryServiceAvailable: isForegroundRecoveryServiceAvailable,
                snapshotBuildResult: snapshotBuildResult,
                paneVisibleSource: paneVisibleSource,
                preContextRecoveryAttempted: preContextRecoveryAttempted,
                preContextRecoverySucceeded: preContextRecoverySucceeded);
        }

        internal static TaskPaneRefreshAttemptResult VisibleAlreadySatisfied()
        {
            return new TaskPaneRefreshAttemptResult(
                true,
                isPaneVisible: true,
                isRefreshCompleted: false,
                completionBasis: "visibleCasePaneAlreadyShown",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.SkippedAlreadyVisible("visibleCasePaneAlreadyShown"),
                paneVisibleSource: PaneVisibleSource.AlreadyVisibleHost);
        }

        internal static TaskPaneRefreshAttemptResult Failed(
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult snapshotBuildResult = null,
            bool preContextRecoveryAttempted = false,
            bool? preContextRecoverySucceeded = null)
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                completionBasis: "failed",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.Unknown("refreshFailed"),
                snapshotBuildResult: snapshotBuildResult,
                preContextRecoveryAttempted: preContextRecoveryAttempted,
                preContextRecoverySucceeded: preContextRecoverySucceeded);
        }

        internal static TaskPaneRefreshAttemptResult Skipped()
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                wasSkipped: true,
                completionBasis: "skipped",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.SkippedNoKnownTarget("refreshSkipped"));
        }

        internal static TaskPaneRefreshAttemptResult ContextRejected(
            bool preContextRecoveryAttempted = false,
            bool? preContextRecoverySucceeded = null)
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                wasContextRejected: true,
                completionBasis: "contextRejected",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.SkippedNoKnownTarget("contextRejected"),
                preContextRecoveryAttempted: preContextRecoveryAttempted,
                preContextRecoverySucceeded: preContextRecoverySucceeded);
        }

        internal TaskPaneRefreshAttemptResult WithForegroundGuaranteeOutcome(ForegroundGuaranteeOutcome foregroundGuaranteeOutcome)
        {
            return new TaskPaneRefreshAttemptResult(
                IsRefreshSucceeded,
                wasSkipped: WasSkipped,
                wasContextRejected: WasContextRejected,
                isPaneVisible: IsPaneVisible,
                isRefreshCompleted: IsRefreshCompleted,
                completionBasis: CompletionBasis,
                foregroundGuaranteeOutcome: foregroundGuaranteeOutcome,
                foregroundContext: ForegroundContext,
                foregroundWorkbook: ForegroundWorkbook,
                foregroundWindow: ForegroundWindow,
                isForegroundRecoveryServiceAvailable: IsForegroundRecoveryServiceAvailable,
                visibilityRecoveryOutcome: VisibilityRecoveryOutcome,
                rebuildFallbackOutcome: RebuildFallbackOutcome,
                refreshSourceSelectionOutcome: RefreshSourceSelectionOutcome,
                snapshotBuildResult: SnapshotBuildResult,
                paneVisibleSource: PaneVisibleSource,
                preContextRecoveryAttempted: PreContextRecoveryAttempted,
                preContextRecoverySucceeded: PreContextRecoverySucceeded);
        }

        internal TaskPaneRefreshAttemptResult WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome visibilityRecoveryOutcome)
        {
            return new TaskPaneRefreshAttemptResult(
                IsRefreshSucceeded,
                wasSkipped: WasSkipped,
                wasContextRejected: WasContextRejected,
                isPaneVisible: IsPaneVisible,
                isRefreshCompleted: IsRefreshCompleted,
                completionBasis: CompletionBasis,
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome,
                foregroundContext: ForegroundContext,
                foregroundWorkbook: ForegroundWorkbook,
                foregroundWindow: ForegroundWindow,
                isForegroundRecoveryServiceAvailable: IsForegroundRecoveryServiceAvailable,
                visibilityRecoveryOutcome: visibilityRecoveryOutcome,
                rebuildFallbackOutcome: RebuildFallbackOutcome,
                refreshSourceSelectionOutcome: RefreshSourceSelectionOutcome,
                snapshotBuildResult: SnapshotBuildResult,
                paneVisibleSource: PaneVisibleSource,
                preContextRecoveryAttempted: PreContextRecoveryAttempted,
                preContextRecoverySucceeded: PreContextRecoverySucceeded);
        }

        internal TaskPaneRefreshAttemptResult WithRebuildFallbackOutcome(RebuildFallbackOutcome rebuildFallbackOutcome)
        {
            return new TaskPaneRefreshAttemptResult(
                IsRefreshSucceeded,
                wasSkipped: WasSkipped,
                wasContextRejected: WasContextRejected,
                isPaneVisible: IsPaneVisible,
                isRefreshCompleted: IsRefreshCompleted,
                completionBasis: CompletionBasis,
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome,
                foregroundContext: ForegroundContext,
                foregroundWorkbook: ForegroundWorkbook,
                foregroundWindow: ForegroundWindow,
                isForegroundRecoveryServiceAvailable: IsForegroundRecoveryServiceAvailable,
                visibilityRecoveryOutcome: VisibilityRecoveryOutcome,
                rebuildFallbackOutcome: rebuildFallbackOutcome,
                refreshSourceSelectionOutcome: RefreshSourceSelectionOutcome,
                snapshotBuildResult: SnapshotBuildResult,
                paneVisibleSource: PaneVisibleSource,
                preContextRecoveryAttempted: PreContextRecoveryAttempted,
                preContextRecoverySucceeded: PreContextRecoverySucceeded);
        }

        internal TaskPaneRefreshAttemptResult WithRefreshSourceSelectionOutcome(RefreshSourceSelectionOutcome refreshSourceSelectionOutcome)
        {
            return new TaskPaneRefreshAttemptResult(
                IsRefreshSucceeded,
                wasSkipped: WasSkipped,
                wasContextRejected: WasContextRejected,
                isPaneVisible: IsPaneVisible,
                isRefreshCompleted: IsRefreshCompleted,
                completionBasis: CompletionBasis,
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome,
                foregroundContext: ForegroundContext,
                foregroundWorkbook: ForegroundWorkbook,
                foregroundWindow: ForegroundWindow,
                isForegroundRecoveryServiceAvailable: IsForegroundRecoveryServiceAvailable,
                visibilityRecoveryOutcome: VisibilityRecoveryOutcome,
                rebuildFallbackOutcome: RebuildFallbackOutcome,
                refreshSourceSelectionOutcome: refreshSourceSelectionOutcome,
                snapshotBuildResult: SnapshotBuildResult,
                paneVisibleSource: PaneVisibleSource,
                preContextRecoveryAttempted: PreContextRecoveryAttempted,
                preContextRecoverySucceeded: PreContextRecoverySucceeded);
        }

        internal bool IsRefreshSucceeded { get; }

        internal bool WasSkipped { get; }

        internal bool WasContextRejected { get; }

        internal bool IsPaneVisible { get; }

        internal bool IsRefreshCompleted { get; }

        internal bool IsForegroundGuaranteeTerminal
        {
            get
            {
                return ForegroundGuaranteeOutcome != null && ForegroundGuaranteeOutcome.IsTerminal;
            }
        }

        internal bool WasForegroundGuaranteeRequired
        {
            get
            {
                return ForegroundGuaranteeOutcome != null && ForegroundGuaranteeOutcome.WasRequired;
            }
        }

        internal string CompletionBasis { get; }

        internal ForegroundGuaranteeOutcome ForegroundGuaranteeOutcome { get; }

        internal WorkbookContext ForegroundContext { get; }

        internal Excel.Workbook ForegroundWorkbook { get; }

        internal Excel.Window ForegroundWindow { get; }

        internal bool IsForegroundRecoveryServiceAvailable { get; }

        internal VisibilityRecoveryOutcome VisibilityRecoveryOutcome { get; }

        internal RebuildFallbackOutcome RebuildFallbackOutcome { get; }

        internal RefreshSourceSelectionOutcome RefreshSourceSelectionOutcome { get; }

        internal TaskPaneSnapshotBuilderService.TaskPaneBuildResult SnapshotBuildResult { get; }

        internal PaneVisibleSource PaneVisibleSource { get; }

        internal bool PreContextRecoveryAttempted { get; }

        internal bool? PreContextRecoverySucceeded { get; }
    }
}
