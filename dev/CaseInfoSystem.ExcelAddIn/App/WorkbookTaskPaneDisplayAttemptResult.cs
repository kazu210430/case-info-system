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
            bool visibleCasePaneAlreadyShown,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts = null)
        {
            AttemptNumber = attemptNumber;
            WorkbookWindow = workbookWindow;
            RefreshAttemptResult = refreshAttemptResult ?? TaskPaneRefreshAttemptResult.Failed();
            VisibleCasePaneAlreadyShown = visibleCasePaneAlreadyShown;
            WorkbookWindowEnsureFacts = workbookWindowEnsureFacts;
        }

        internal int AttemptNumber { get; }

        internal Excel.Window WorkbookWindow { get; }

        internal TaskPaneRefreshAttemptResult RefreshAttemptResult { get; }

        internal bool VisibleCasePaneAlreadyShown { get; }

        internal WorkbookWindowVisibilityEnsureFacts WorkbookWindowEnsureFacts { get; }

        internal bool IsShown
        {
            get
            {
                return RefreshAttemptResult.IsRefreshSucceeded && RefreshAttemptResult.IsPaneVisible;
            }
        }

        internal WorkbookTaskPaneReadyShowAttemptOutcome WithWorkbookWindowEnsureFacts(WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            return new WorkbookTaskPaneReadyShowAttemptOutcome(
                AttemptNumber,
                WorkbookWindow,
                RefreshAttemptResult,
                VisibleCasePaneAlreadyShown,
                workbookWindowEnsureFacts);
        }
    }

    internal sealed class WorkbookWindowVisibilityEnsureFacts
    {
        private WorkbookWindowVisibilityEnsureFacts(
            WorkbookWindowVisibilityEnsureOutcome outcome,
            string workbookFullName,
            string windowHwnd,
            bool? visibleAfterSet)
        {
            Outcome = outcome;
            WorkbookFullName = workbookFullName ?? string.Empty;
            WindowHwnd = windowHwnd ?? string.Empty;
            VisibleAfterSet = visibleAfterSet;
        }

        internal WorkbookWindowVisibilityEnsureOutcome Outcome { get; }

        internal string WorkbookFullName { get; }

        internal string WindowHwnd { get; }

        internal bool? VisibleAfterSet { get; }

        internal static WorkbookWindowVisibilityEnsureFacts FromResult(WorkbookWindowVisibilityEnsureResult result)
        {
            if (result == null)
            {
                return null;
            }

            return new WorkbookWindowVisibilityEnsureFacts(
                result.Outcome,
                result.WorkbookFullName,
                result.WindowHwnd,
                result.VisibleAfterSet);
        }
    }
}
