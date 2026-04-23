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

    internal static class PaneDisplayPolicy
    {
        internal static PaneDisplayPolicyResult Decide(
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

        internal static PaneDisplayPolicyResult Decide(
            TaskPaneDisplayRequest request,
            TaskPaneManager taskPaneManager,
            Excel.Workbook workbook,
            Excel.Window window,
            bool shouldDisplayPane)
        {
            if (!IsAcceptedRequest(request))
            {
                return PaneDisplayPolicyResult.Reject;
            }

            if (window == null)
            {
                return PaneDisplayPolicyResult.Reject;
            }

            if (!TryResolveWindowKey(window, out _))
            {
                return PaneDisplayPolicyResult.Reject;
            }

            if (taskPaneManager == null)
            {
                return shouldDisplayPane
                    ? PaneDisplayPolicyResult.ShowWithRender
                    : PaneDisplayPolicyResult.Reject;
            }

            bool hasManagedPane = taskPaneManager.HasManagedPaneForWindow(window);
            bool showedExistingPane = false;
            bool shouldShowWithRenderPane = false;
            if (shouldDisplayPane)
            {
                showedExistingPane = taskPaneManager.TryShowExistingPaneForDisplayRequest(workbook, window);
                shouldShowWithRenderPane = !showedExistingPane
                    && taskPaneManager.ShouldShowWithRenderPaneForDisplayRequest(workbook, window);
            }

            return Decide(
                shouldDisplayPane,
                hasManagedPane,
                showedExistingPane,
                shouldShowWithRenderPane);
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
