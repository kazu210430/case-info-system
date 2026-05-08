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
            bool isForegroundRecoveryServiceAvailable = false)
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
                    : ForegroundGuaranteeOutcome.NotRequired(completionBasis));
        }

        internal static TaskPaneRefreshAttemptResult RefreshCompletedPendingForeground(
            WorkbookContext foregroundContext,
            Excel.Workbook foregroundWorkbook,
            Excel.Window foregroundWindow,
            bool isForegroundRecoveryServiceAvailable,
            string completionBasis)
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
                isForegroundRecoveryServiceAvailable: isForegroundRecoveryServiceAvailable);
        }

        internal static TaskPaneRefreshAttemptResult VisibleAlreadySatisfied()
        {
            return new TaskPaneRefreshAttemptResult(
                true,
                isPaneVisible: true,
                isRefreshCompleted: false,
                completionBasis: "visibleCasePaneAlreadyShown",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.SkippedAlreadyVisible("visibleCasePaneAlreadyShown"));
        }

        internal static TaskPaneRefreshAttemptResult Failed()
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                completionBasis: "failed",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.Unknown("refreshFailed"));
        }

        internal static TaskPaneRefreshAttemptResult Skipped()
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                wasSkipped: true,
                completionBasis: "skipped",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.SkippedNoKnownTarget("refreshSkipped"));
        }

        internal static TaskPaneRefreshAttemptResult ContextRejected()
        {
            return new TaskPaneRefreshAttemptResult(
                false,
                wasContextRejected: true,
                completionBasis: "contextRejected",
                foregroundGuaranteeOutcome: ForegroundGuaranteeOutcome.SkippedNoKnownTarget("contextRejected"));
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
                isForegroundRecoveryServiceAvailable: IsForegroundRecoveryServiceAvailable);
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
    }
}
