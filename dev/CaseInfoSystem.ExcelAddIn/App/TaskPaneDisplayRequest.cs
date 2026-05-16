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
            string detail,
            WindowActivateTaskPaneTriggerFacts windowActivateTriggerFacts = null)
        {
            Source = source;
            RefreshIntent = refreshIntent;
            Detail = detail ?? string.Empty;
            WindowActivateTriggerFacts = windowActivateTriggerFacts;
        }

        internal TaskPaneDisplaySource Source { get; }

        internal TaskPaneDisplayRefreshIntent RefreshIntent { get; }

        internal string Detail { get; }

        internal WindowActivateTaskPaneTriggerFacts WindowActivateTriggerFacts { get; }

        internal static TaskPaneDisplayRequest ForWindowActivate(WindowActivateTaskPaneTriggerFacts triggerFacts = null)
        {
            return new TaskPaneDisplayRequest(
                TaskPaneDisplaySource.WindowActivate,
                TaskPaneDisplayRefreshIntent.Normal,
                string.Empty,
                triggerFacts);
        }

        internal static TaskPaneDisplayRequest ForPostActionRefresh(string actionKind)
        {
            return new TaskPaneDisplayRequest(
                TaskPaneDisplaySource.PostActionRefresh,
                TaskPaneDisplayRefreshIntent.ForceRefresh,
                actionKind,
                windowActivateTriggerFacts: null);
        }

        internal bool IsWindowActivateTrigger
        {
            get { return Source == TaskPaneDisplaySource.WindowActivate; }
        }

        internal string ToReasonString()
        {
            if (Source == TaskPaneDisplaySource.WindowActivate)
            {
                return ControlFlowReasons.WindowActivate;
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
