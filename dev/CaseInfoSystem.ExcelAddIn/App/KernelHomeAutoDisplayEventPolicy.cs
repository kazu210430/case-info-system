using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelHomeAutoDisplayEventPolicy
    {
        internal static bool ShouldAutoShow(string eventName, bool startupPolicyAllowsDisplay)
        {
            if (!startupPolicyAllowsDisplay)
            {
                return false;
            }

            return string.Equals(eventName, "WorkbookOpen", StringComparison.OrdinalIgnoreCase);
        }
    }
}
