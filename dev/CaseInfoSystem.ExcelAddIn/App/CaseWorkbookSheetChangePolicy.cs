namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum CaseWorkbookSheetChangeAction
    {
        Ignore = 0,
        IgnoreBecauseTransientPaneSuppression = 1,
        MarkSessionDirty = 2
    }

    internal static class CaseWorkbookSheetChangePolicy
    {
        internal static CaseWorkbookSheetChangeAction Decide(
            bool isBaseOrCaseWorkbook,
            bool isManagedClose,
            bool isTransientPaneSuppressed)
        {
            if (!isBaseOrCaseWorkbook || isManagedClose)
            {
                return CaseWorkbookSheetChangeAction.Ignore;
            }

            return isTransientPaneSuppressed
                ? CaseWorkbookSheetChangeAction.IgnoreBecauseTransientPaneSuppression
                : CaseWorkbookSheetChangeAction.MarkSessionDirty;
        }
    }
}
