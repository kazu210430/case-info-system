using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class TaskPaneRefreshPreconditionPolicy
    {
        internal static bool ShouldHideAllAndSkip(WorkbookRole role, string windowKey)
        {
            if (role == WorkbookRole.Unknown)
            {
                return true;
            }

            return windowKey != null && string.IsNullOrWhiteSpace(windowKey);
        }
    }
}
