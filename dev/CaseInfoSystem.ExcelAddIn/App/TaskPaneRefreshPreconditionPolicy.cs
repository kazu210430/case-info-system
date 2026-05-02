using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class TaskPaneRefreshPreconditionPolicy
    {
        internal static bool ShouldSkipWorkbookOpenWindowDependentRefresh(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return string.Equals(reason, "WorkbookOpen", StringComparison.Ordinal)
                && workbook != null
                && window == null;
        }

        internal static bool ShouldHideAllAndSkip(WorkbookRole role, string windowKey)
        {
            if (role == WorkbookRole.Unknown)
            {
                return true;
            }

            return windowKey != null && string.IsNullOrWhiteSpace(windowKey);
        }
    }
}
