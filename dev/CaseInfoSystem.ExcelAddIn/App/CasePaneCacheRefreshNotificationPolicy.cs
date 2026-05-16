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

            return string.Equals(reason, ControlFlowReasons.WorkbookOpen, StringComparison.OrdinalIgnoreCase)
                || string.Equals(reason, ControlFlowReasons.WorkbookActivate, StringComparison.OrdinalIgnoreCase);
        }
    }
}
