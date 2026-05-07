using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum PaneDisplayPolicyResult
    {
        ShowExisting = 0,
        ShowWithRender = 1,
        Reject = 2,
        Hide = 3
    }

    internal sealed class TaskPaneDisplayEntryDecision
    {
        internal TaskPaneDisplayEntryDecision(PaneDisplayPolicyResult result, TaskPaneDisplayEntryState state)
        {
            Result = result;
            State = state;
        }

        internal PaneDisplayPolicyResult Result { get; }

        internal TaskPaneDisplayEntryState State { get; }
    }

    internal static class PaneDisplayPolicy
    {
        internal static TaskPaneDisplayEntryDecision Decide(
            TaskPaneDisplayRequest request,
            TaskPaneManager taskPaneManager,
            IWorkbookRoleResolver workbookRoleResolver,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            return Decide(
                request,
                taskPaneManager,
                workbook,
                window,
                ShouldDisplayPane(workbookRoleResolver, workbook));
        }

        internal static TaskPaneDisplayEntryDecision Decide(
            TaskPaneDisplayRequest request,
            TaskPaneManager taskPaneManager,
            Excel.Workbook workbook,
            Excel.Window window,
            bool shouldDisplayPane)
        {
            if (!IsAcceptedRequest(request))
            {
                return new TaskPaneDisplayEntryDecision(PaneDisplayPolicyResult.Reject, state: null);
            }

            if (taskPaneManager == null)
            {
                if (window == null)
                {
                    return new TaskPaneDisplayEntryDecision(PaneDisplayPolicyResult.Reject, state: null);
                }

                if (!TryResolveWindowKey(window, out _))
                {
                    return new TaskPaneDisplayEntryDecision(PaneDisplayPolicyResult.Reject, state: null);
                }

                return new TaskPaneDisplayEntryDecision(
                    shouldDisplayPane
                        ? PaneDisplayPolicyResult.ShowWithRender
                        : PaneDisplayPolicyResult.Reject,
                    state: null);
            }

            TaskPaneDisplayEntryState state = taskPaneManager.EvaluateDisplayEntryState(workbook, window);
            return Decide(request, state, shouldDisplayPane);
        }

        internal static TaskPaneDisplayEntryDecision Decide(
            TaskPaneDisplayRequest request,
            TaskPaneDisplayEntryState state,
            bool shouldDisplayPane)
        {
            if (!IsAcceptedRequest(request)
                || state == null
                || !state.HasTargetWindow
                || !state.HasResolvableWindowKey)
            {
                return new TaskPaneDisplayEntryDecision(PaneDisplayPolicyResult.Reject, state);
            }

            bool shouldShowExisting = shouldDisplayPane && TaskPaneShowExistingPolicy.ShouldShowExisting(
                hasExistingHost: state.HasExistingHost,
                isSameWorkbook: state.IsSameWorkbook,
                isRenderSignatureCurrent: state.IsRenderSignatureCurrent);
            bool shouldShowWithRenderPane = shouldDisplayPane && !shouldShowExisting && TaskPaneShowWithRenderPolicy.ShouldShowWithRender(
                state.HasExistingHost,
                state.IsSameWorkbook,
                state.IsRenderSignatureCurrent);

            return new TaskPaneDisplayEntryDecision(
                Decide(
                    shouldDisplayPane,
                    state.HasManagedPane,
                    shouldShowExisting,
                    shouldShowWithRenderPane),
                state);
        }

        internal static bool ShouldDisplayPane(IWorkbookRoleResolver workbookRoleResolver, Excel.Workbook workbook)
        {
            if (workbookRoleResolver == null)
            {
                return true;
            }

            WorkbookRole role = workbookRoleResolver.Resolve(workbook);
            return role == WorkbookRole.Kernel
                || role == WorkbookRole.Case
                || role == WorkbookRole.Accounting;
        }

        internal static PaneDisplayPolicyResult Decide(bool showedExistingPane, bool shouldShowWithRenderPane)
        {
            return Decide(
                shouldDisplayPane: true,
                hasManagedPane: showedExistingPane,
                showedExistingPane: showedExistingPane,
                shouldShowWithRenderPane: shouldShowWithRenderPane);
        }

        internal static PaneDisplayPolicyResult Decide(
            bool shouldDisplayPane,
            bool hasManagedPane,
            bool showedExistingPane,
            bool shouldShowWithRenderPane)
        {
            if (showedExistingPane)
            {
                return PaneDisplayPolicyResult.ShowExisting;
            }

            if (!shouldDisplayPane)
            {
                return hasManagedPane
                    ? PaneDisplayPolicyResult.Hide
                    : PaneDisplayPolicyResult.Reject;
            }

            if (shouldShowWithRenderPane)
            {
                return PaneDisplayPolicyResult.ShowWithRender;
            }

            if (TaskPaneDisplayRejectPolicy.ShouldReject(showedExistingPane, shouldShowWithRenderPane))
            {
                return PaneDisplayPolicyResult.Reject;
            }

            return PaneDisplayPolicyResult.ShowWithRender;
        }

        private static bool IsAcceptedRequest(TaskPaneDisplayRequest request)
        {
            return request != null;
        }

        private static bool TryResolveWindowKey(Excel.Window window, out string windowKey)
        {
            windowKey = string.Empty;

            if (window == null)
            {
                return false;
            }

            try
            {
                windowKey = window.Hwnd.ToString();
                return !string.IsNullOrWhiteSpace(windowKey);
            }
            catch
            {
                return false;
            }
        }
    }
}
