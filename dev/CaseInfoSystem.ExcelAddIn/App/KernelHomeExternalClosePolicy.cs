namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelHomeExternalClosePolicy
    {
        internal static bool ShouldCloseKernelHome(bool isKernelCaseCreationFlowActive)
        {
            return isKernelCaseCreationFlowActive;
        }
    }
}
