using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookTaskPaneDisplayAttemptResult
    {
        internal WorkbookTaskPaneDisplayAttemptResult(Excel.Window workbookWindow, TaskPaneRefreshAttemptResult refreshAttemptResult)
        {
            WorkbookWindow = workbookWindow;
            RefreshAttemptResult = refreshAttemptResult;
        }

        internal Excel.Window WorkbookWindow { get; }

        internal TaskPaneRefreshAttemptResult RefreshAttemptResult { get; }
    }
}
