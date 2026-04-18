namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum CaseWorkbookBeforeCloseAction
    {
        Ignore = 0,
        SuppressPromptForManagedClose = 1,
        PromptForDirtySession = 2,
        SchedulePostCloseFollowUp = 3
    }

    internal static class CaseWorkbookBeforeClosePolicy
    {
        internal static CaseWorkbookBeforeCloseAction Decide(
            bool isBaseOrCaseWorkbook,
            bool isManagedClose,
            bool isSessionDirty)
        {
            if (!isBaseOrCaseWorkbook)
            {
                return CaseWorkbookBeforeCloseAction.Ignore;
            }

            if (isManagedClose)
            {
                return CaseWorkbookBeforeCloseAction.SuppressPromptForManagedClose;
            }

            return isSessionDirty
                ? CaseWorkbookBeforeCloseAction.PromptForDirtySession
                : CaseWorkbookBeforeCloseAction.SchedulePostCloseFollowUp;
        }
    }
}
