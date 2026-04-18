namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelWorkbookStartupDisplayPolicy
    {
        internal static bool ShouldShowHomeOnStartup(
            bool hasExplicitKernelStartupContext,
            bool hasKernelWorkbookContext,
            bool isStartupWorkbookKernel,
            bool hasVisibleNonKernelWorkbook)
        {
            if (!hasExplicitKernelStartupContext)
            {
                return false;
            }

            if (!hasKernelWorkbookContext)
            {
                return false;
            }

            if (isStartupWorkbookKernel)
            {
                return true;
            }

            return !hasVisibleNonKernelWorkbook;
        }
    }
}
