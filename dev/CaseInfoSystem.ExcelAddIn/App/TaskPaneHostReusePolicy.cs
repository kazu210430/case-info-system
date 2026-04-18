using System;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class TaskPaneHostReusePolicy
    {
        internal static bool ShouldReuseCaseHostWithoutRender(
            WorkbookRole role,
            bool isDocumentButtonsHost,
            bool isAlreadyRendered,
            bool isSameWorkbook,
            string reason)
        {
            if (role != WorkbookRole.Case || !isDocumentButtonsHost)
            {
                return false;
            }

            if (!isAlreadyRendered || !isSameWorkbook)
            {
                return false;
            }

            return string.Equals(reason, "WorkbookActivate", StringComparison.OrdinalIgnoreCase)
                || string.Equals(reason, "WindowActivate", StringComparison.OrdinalIgnoreCase)
                || string.Equals(reason, "KernelHomeForm.FormClosed", StringComparison.OrdinalIgnoreCase);
        }
    }
}
