namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelWorkbookWindowRestorePolicy
    {
        internal static bool ShouldAvoidGlobalExcelWindowRestore(
            bool isKernelCaseCreationFlowActive,
            bool hasVisibleNonKernelWorkbook)
        {
            return !isKernelCaseCreationFlowActive && hasVisibleNonKernelWorkbook;
        }
    }
}
