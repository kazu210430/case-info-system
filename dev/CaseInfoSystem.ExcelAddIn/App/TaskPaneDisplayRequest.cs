namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum TaskPaneDisplaySource
    {
        WindowActivate = 0,
        PostActionRefresh = 1
    }

    internal enum TaskPaneDisplayRefreshIntent
    {
        Normal = 0,
        ForceRefresh = 1
    }

    internal sealed class TaskPaneDisplayRequest
    {
        internal TaskPaneDisplayRequest(
            TaskPaneDisplaySource source,
            TaskPaneDisplayRefreshIntent refreshIntent,
            string detail)
        {
            Source = source;
            RefreshIntent = refreshIntent;
            Detail = detail ?? string.Empty;
        }

        internal TaskPaneDisplaySource Source { get; }

        internal TaskPaneDisplayRefreshIntent RefreshIntent { get; }

        internal string Detail { get; }

        internal static TaskPaneDisplayRequest ForWindowActivate()
        {
            return new TaskPaneDisplayRequest(
                TaskPaneDisplaySource.WindowActivate,
                TaskPaneDisplayRefreshIntent.Normal,
                string.Empty);
        }

        internal static TaskPaneDisplayRequest ForPostActionRefresh(string actionKind)
        {
            return new TaskPaneDisplayRequest(
                TaskPaneDisplaySource.PostActionRefresh,
                TaskPaneDisplayRefreshIntent.ForceRefresh,
                actionKind);
        }

        internal string ToReasonString()
        {
            if (Source == TaskPaneDisplaySource.WindowActivate)
            {
                return "WindowActivate";
            }

            return string.IsNullOrWhiteSpace(Detail)
                ? "PostActionRefresh"
                : "PostActionRefresh." + Detail;
        }
    }

    internal static class TaskPaneDisplayRejectPolicy
    {
        internal static bool ShouldReject(bool showedExistingPane, bool shouldShowWithRenderPane)
        {
            return !showedExistingPane && !shouldShowWithRenderPane;
        }
    }
}
