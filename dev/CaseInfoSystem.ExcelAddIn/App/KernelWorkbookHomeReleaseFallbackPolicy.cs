namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum KernelWorkbookHomeReleaseAction
    {
        SkipRestore = 0,
        RestoreWithoutPromotion = 1,
        PromoteAndRestore = 2
    }

    internal static class KernelWorkbookHomeReleaseFallbackPolicy
    {
        internal static KernelWorkbookHomeReleaseAction DecideHomeReleaseAction(
            bool shouldAvoidGlobalExcelWindowRestore,
            bool shouldPromoteKernelWorkbook)
        {
            if (shouldAvoidGlobalExcelWindowRestore)
            {
                return KernelWorkbookHomeReleaseAction.SkipRestore;
            }

            if (shouldPromoteKernelWorkbook)
            {
                return KernelWorkbookHomeReleaseAction.PromoteAndRestore;
            }

            return KernelWorkbookHomeReleaseAction.RestoreWithoutPromotion;
        }
    }
}
