namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum TaskPaneHostPreparationAction
    {
        None = 0,
        HideAllExceptActiveWindow = 1,
        HideNonCaseHostsExceptActiveWindow = 2
    }

    internal static class TaskPaneHostPreparationPolicy
    {
        internal static TaskPaneHostPreparationAction Decide(bool isKernelCaseCreationFlowActive, bool isCaseHost)
        {
            if (!isKernelCaseCreationFlowActive)
            {
                return TaskPaneHostPreparationAction.None;
            }

            return isCaseHost
                ? TaskPaneHostPreparationAction.HideNonCaseHostsExceptActiveWindow
                : TaskPaneHostPreparationAction.HideAllExceptActiveWindow;
        }
    }
}
