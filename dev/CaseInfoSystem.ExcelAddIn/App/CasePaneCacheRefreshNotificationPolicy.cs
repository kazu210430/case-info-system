using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class CasePaneCacheRefreshNotificationPolicy
    {
        internal static bool ShouldNotify(bool updatedCaseSnapshotCache, string reason)
        {
            if (!updatedCaseSnapshotCache)
            {
                return false;
            }

            return string.Equals(reason, "WorkbookOpen", StringComparison.OrdinalIgnoreCase)
                || string.Equals(reason, "WorkbookActivate", StringComparison.OrdinalIgnoreCase);
        }
    }
}
