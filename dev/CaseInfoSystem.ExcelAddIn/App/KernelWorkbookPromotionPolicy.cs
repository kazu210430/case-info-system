namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelWorkbookPromotionPolicy
    {
        internal static bool ShouldPromoteKernelWorkbookOnHomeRelease(
            bool isKernelCaseCreationFlowActive,
            bool hasActiveWorkbook,
            bool isActiveWorkbookKernel,
            bool hasVisibleNonKernelWorkbook)
        {
            if (isKernelCaseCreationFlowActive)
            {
                return true;
            }

            if (!hasActiveWorkbook)
            {
                return !hasVisibleNonKernelWorkbook;
            }

            if (isActiveWorkbookKernel)
            {
                return true;
            }

            return false;
        }
    }
}
