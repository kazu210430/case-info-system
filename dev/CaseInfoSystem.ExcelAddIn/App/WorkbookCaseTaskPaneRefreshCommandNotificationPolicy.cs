namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum WorkbookCaseTaskPaneRefreshCommandNotificationKind
    {
        Failed = 0,
        Updated = 1,
        Latest = 2
    }

    internal static class WorkbookCaseTaskPaneRefreshCommandNotificationPolicy
    {
        internal static WorkbookCaseTaskPaneRefreshCommandNotificationKind Decide(TaskPaneRefreshAttemptResult refreshResult)
        {
            if (refreshResult == null)
            {
                return WorkbookCaseTaskPaneRefreshCommandNotificationKind.Failed;
            }

            if (!refreshResult.IsRefreshSucceeded)
            {
                if (refreshResult.WasSkipped
                    && string.Equals(refreshResult.CompletionBasis, "ignore-during-protection", System.StringComparison.OrdinalIgnoreCase))
                {
                    return WorkbookCaseTaskPaneRefreshCommandNotificationKind.Latest;
                }

                return WorkbookCaseTaskPaneRefreshCommandNotificationKind.Failed;
            }

            if (refreshResult.SnapshotBuildResult != null
                && refreshResult.SnapshotBuildResult.UpdatedCaseSnapshotCache)
            {
                return WorkbookCaseTaskPaneRefreshCommandNotificationKind.Updated;
            }

            return WorkbookCaseTaskPaneRefreshCommandNotificationKind.Latest;
        }
    }
}
