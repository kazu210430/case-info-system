using CaseInfoSystem.ExcelAddIn.Domain;
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
            bool preContextRecoveryAttempted = false,
            bool? preContextRecoverySucceeded = null)
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                completionBasis: "failed",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.Unknown("refreshFailed"),
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

        internal PaneVisibleSource PaneVisibleSource { get; }

        internal bool PreContextRecoveryAttempted { get; }

        internal bool? PreContextRecoverySucceeded { get; }
    }
}
