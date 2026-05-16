using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum TaskPaneHostFlowPreconditionDecision
    {
        Proceed = 0,
        HideAllAndSkipForUnknownRole = 1,
        HideAllAndSkipForMissingWindowKey = 2
    }

    internal sealed class TaskPaneRefreshPreconditionDecision
    {
        private TaskPaneRefreshPreconditionDecision(bool canRefresh, string skipActionName)
        {
            CanRefresh = canRefresh;
            SkipActionName = skipActionName ?? string.Empty;
        }

        internal bool CanRefresh { get; }

        internal string SkipActionName { get; }

        internal static TaskPaneRefreshPreconditionDecision Proceed()
        {
            return new TaskPaneRefreshPreconditionDecision(true, string.Empty);
        }

        internal static TaskPaneRefreshPreconditionDecision SkipWorkbookOpenWindowDependentRefresh()
        {
            return new TaskPaneRefreshPreconditionDecision(false, "skip-workbook-open-window-dependent-refresh");
        }

        internal static TaskPaneRefreshPreconditionDecision IgnoreDuringProtection()
        {
            return new TaskPaneRefreshPreconditionDecision(false, "ignore-during-protection");
        }
    }

    internal static class TaskPaneRefreshPreconditionPolicy
    {
        internal static TaskPaneRefreshPreconditionDecision DecideRefreshPrecondition(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            Func<bool> shouldIgnoreDuringProtection)
        {
            if (ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window))
            {
                return TaskPaneRefreshPreconditionDecision.SkipWorkbookOpenWindowDependentRefresh();
            }

            if (shouldIgnoreDuringProtection != null && shouldIgnoreDuringProtection())
            {
                return TaskPaneRefreshPreconditionDecision.IgnoreDuringProtection();
            }

            return TaskPaneRefreshPreconditionDecision.Proceed();
        }

        internal static bool ShouldSkipWorkbookOpenWindowDependentRefresh(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return string.Equals(reason, ControlFlowReasons.WorkbookOpen, StringComparison.Ordinal)
                && workbook != null
                && window == null;
        }

        internal static TaskPaneHostFlowPreconditionDecision DecideHostFlowPrecondition(WorkbookRole role, string windowKey)
        {
            if (role == WorkbookRole.Unknown)
            {
                return TaskPaneHostFlowPreconditionDecision.HideAllAndSkipForUnknownRole;
            }

            if (windowKey != null && string.IsNullOrWhiteSpace(windowKey))
            {
                return TaskPaneHostFlowPreconditionDecision.HideAllAndSkipForMissingWindowKey;
            }

            return TaskPaneHostFlowPreconditionDecision.Proceed;
        }
    }
}
