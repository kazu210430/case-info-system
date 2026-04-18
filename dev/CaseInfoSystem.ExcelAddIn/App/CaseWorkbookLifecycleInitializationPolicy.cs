namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum CaseWorkbookInitializationAction
    {
        None = 0,
        InitializeBaseWorkbook = 1,
        InitializeCaseWorkbook = 2
    }

    internal static class CaseWorkbookLifecycleInitializationPolicy
    {
        internal static CaseWorkbookInitializationAction Decide(
            bool isBaseOrCaseWorkbook,
            string workbookKey,
            bool isAlreadyInitialized,
            bool isCaseWorkbook)
        {
            if (!isBaseOrCaseWorkbook || string.IsNullOrWhiteSpace(workbookKey) || isAlreadyInitialized)
            {
                return CaseWorkbookInitializationAction.None;
            }

            return isCaseWorkbook
                ? CaseWorkbookInitializationAction.InitializeCaseWorkbook
                : CaseWorkbookInitializationAction.InitializeBaseWorkbook;
        }
    }
}
