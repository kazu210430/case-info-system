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

    internal sealed class WorkbookTaskPaneReadyShowAttemptOutcome
    {
        internal WorkbookTaskPaneReadyShowAttemptOutcome(
            int attemptNumber,
            Excel.Window workbookWindow,
            TaskPaneRefreshAttemptResult refreshAttemptResult,
            bool visibleCasePaneAlreadyShown)
        {
            AttemptNumber = attemptNumber;
            WorkbookWindow = workbookWindow;
            RefreshAttemptResult = refreshAttemptResult ?? TaskPaneRefreshAttemptResult.Failed();
            VisibleCasePaneAlreadyShown = visibleCasePaneAlreadyShown;
        }

        internal int AttemptNumber { get; }

        internal Excel.Window WorkbookWindow { get; }

        internal TaskPaneRefreshAttemptResult RefreshAttemptResult { get; }

        internal bool VisibleCasePaneAlreadyShown { get; }

        internal bool IsShown
        {
            get
            {
                return RefreshAttemptResult.IsRefreshSucceeded && RefreshAttemptResult.IsPaneVisible;
            }
        }
    }
}
