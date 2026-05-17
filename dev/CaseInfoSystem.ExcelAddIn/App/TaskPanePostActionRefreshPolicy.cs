using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum TaskPanePostActionRefreshDecision
    {
        RefreshImmediately,
        SkipForForegroundPreservation,
        DeferAndInvalidateSignature
    }

    internal static class TaskPanePostActionRefreshPolicy
    {
        internal static TaskPanePostActionRefreshDecision Decide(string actionKind)
        {
            if (string.Equals(actionKind, "doc", StringComparison.OrdinalIgnoreCase))
            {
                return TaskPanePostActionRefreshDecision.SkipForForegroundPreservation;
            }

            if (string.Equals(actionKind, "accounting", StringComparison.OrdinalIgnoreCase))
            {
                return TaskPanePostActionRefreshDecision.SkipForForegroundPreservation;
            }

            if (string.Equals(actionKind, "caselist", StringComparison.OrdinalIgnoreCase))
            {
                return TaskPanePostActionRefreshDecision.DeferAndInvalidateSignature;
            }

            return TaskPanePostActionRefreshDecision.RefreshImmediately;
        }
    }
}
