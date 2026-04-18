using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookTaskPaneDisplayAttemptCoordinator
    {
        internal WorkbookTaskPaneDisplayAttemptResult TryShowOnce(
            Excel.Workbook workbook,
            string reason,
            Func<Excel.Workbook, string, Excel.Window> resolveWorkbookPaneWindow,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane)
        {
            Excel.Window workbookWindow = resolveWorkbookPaneWindow(workbook, reason);
            TaskPaneRefreshAttemptResult refreshAttemptResult = tryRefreshTaskPane(reason, workbook, workbookWindow);
            return new WorkbookTaskPaneDisplayAttemptResult(workbookWindow, refreshAttemptResult);
        }
    }
}
