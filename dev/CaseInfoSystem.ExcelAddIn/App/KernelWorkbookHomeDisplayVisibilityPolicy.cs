namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum KernelWorkbookHomeDisplayVisibilityAction
    {
        MinimizeKernelWindows = 0,
        ConcealKernelWindowsAndHideExcelMainWindow = 1,
        HideExcelMainWindowOnly = 2
    }

    internal static class KernelWorkbookHomeDisplayVisibilityPolicy
    {
        internal static KernelWorkbookHomeDisplayVisibilityAction DecideAction(
            bool hasVisibleNonKernelWorkbook,
            bool isActiveWorkbookKernel,
            int visibleWorkbookCount)
        {
            if (hasVisibleNonKernelWorkbook)
            {
                return KernelWorkbookHomeDisplayVisibilityAction.MinimizeKernelWindows;
            }

            if (isActiveWorkbookKernel && visibleWorkbookCount >= 1)
            {
                return KernelWorkbookHomeDisplayVisibilityAction.ConcealKernelWindowsAndHideExcelMainWindow;
            }

            return KernelWorkbookHomeDisplayVisibilityAction.HideExcelMainWindowOnly;
        }
    }
}
