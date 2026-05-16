using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshPreconditionDecisionService
    {
        internal TaskPaneRefreshPreconditionDecisionResult Decide(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            Func<bool> shouldIgnoreDuringProtection)
        {
            TaskPaneRefreshPreconditionDecision preconditionDecision = TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(
                reason,
                workbook,
                window,
                shouldIgnoreDuringProtection);
            if (preconditionDecision.CanRefresh)
            {
                return TaskPaneRefreshPreconditionDecisionResult.Continue(preconditionDecision);
            }

            return TaskPaneRefreshPreconditionDecisionResult.FailClosed(
                preconditionDecision,
                TaskPaneRefreshFailClosedOutcome.FromPreconditionDecision(preconditionDecision));
        }
    }

    internal sealed class TaskPaneRefreshPreconditionDecisionResult
    {
        private TaskPaneRefreshPreconditionDecisionResult(
            bool canRefresh,
            TaskPaneRefreshPreconditionDecision preconditionDecision,
            TaskPaneRefreshFailClosedOutcome? failClosedOutcome)
        {
            CanRefresh = canRefresh;
            PreconditionDecision = preconditionDecision;
            FailClosedOutcome = failClosedOutcome;
        }

        internal bool CanRefresh { get; }

        internal bool ShouldFailClosed
        {
            get { return !CanRefresh; }
        }

        internal TaskPaneRefreshPreconditionDecision PreconditionDecision { get; }

        internal TaskPaneRefreshFailClosedOutcome? FailClosedOutcome { get; }

        internal string SkipReason
        {
            get { return PreconditionDecision == null ? string.Empty : PreconditionDecision.SkipActionName; }
        }

        internal string SkipActionName
        {
            get { return SkipReason; }
        }

        internal TaskPaneRefreshAttemptResult NormalizedOutcomeAttemptResult
        {
            get { return FailClosedOutcome.HasValue ? FailClosedOutcome.Value.AttemptResult : null; }
        }

        internal string NormalizedOutcomeActionName
        {
            get { return FailClosedOutcome.HasValue ? FailClosedOutcome.Value.SkipActionName : string.Empty; }
        }

        internal static TaskPaneRefreshPreconditionDecisionResult Continue(
            TaskPaneRefreshPreconditionDecision preconditionDecision)
        {
            return new TaskPaneRefreshPreconditionDecisionResult(
                canRefresh: true,
                preconditionDecision: preconditionDecision,
                failClosedOutcome: null);
        }

        internal static TaskPaneRefreshPreconditionDecisionResult FailClosed(
            TaskPaneRefreshPreconditionDecision preconditionDecision,
            TaskPaneRefreshFailClosedOutcome failClosedOutcome)
        {
            return new TaskPaneRefreshPreconditionDecisionResult(
                canRefresh: false,
                preconditionDecision: preconditionDecision,
                failClosedOutcome: failClosedOutcome);
        }
    }
}
