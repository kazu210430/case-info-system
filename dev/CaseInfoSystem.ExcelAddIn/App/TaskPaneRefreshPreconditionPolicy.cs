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

    internal static class TaskPaneRefreshPreconditionPolicy
    {
        internal static bool ShouldSkipWorkbookOpenWindowDependentRefresh(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return string.Equals(reason, "WorkbookOpen", StringComparison.Ordinal)
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
