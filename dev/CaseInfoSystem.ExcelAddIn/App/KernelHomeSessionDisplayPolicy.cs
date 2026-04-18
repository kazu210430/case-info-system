namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum KernelHomeSessionCompletionAction
    {
        ReleaseHomeDisplayWithShowingExcel = 0,
        ReleaseHomeDisplayWithoutShowingExcelAndQuit = 1,
        DismissPreparedHomeDisplayState = 2
    }

    internal static class KernelHomeSessionDisplayPolicy
    {
        internal static bool ShouldSkipDisplayRestoreForCaseCreation(
            bool saveKernelWorkbook,
            bool isKernelCaseCreationFlowActive,
            bool otherVisibleWorkbookExists,
            bool otherWorkbookExists)
        {
            return saveKernelWorkbook
                && isKernelCaseCreationFlowActive
                && (otherVisibleWorkbookExists || otherWorkbookExists);
        }

        internal static KernelHomeSessionCompletionAction DecideCompletionAction(
            bool skipDisplayRestoreForCaseCreation,
            bool otherVisibleWorkbookExists,
            bool otherWorkbookExists)
        {
            if (!otherVisibleWorkbookExists && !otherWorkbookExists)
            {
                return KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithoutShowingExcelAndQuit;
            }

            return skipDisplayRestoreForCaseCreation
                ? KernelHomeSessionCompletionAction.DismissPreparedHomeDisplayState
                : KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithShowingExcel;
        }
    }
}
