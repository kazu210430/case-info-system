using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum PaneDisplayPolicyResult
    {
        ShowExisting = 0,
        ShowWithRender = 1,
        Reject = 2
    }

    internal static class PaneDisplayPolicy
    {
        internal static PaneDisplayPolicyResult Decide(
            TaskPaneManager taskPaneManager,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            if (taskPaneManager == null)
            {
                return PaneDisplayPolicyResult.ShowWithRender;
            }

            bool showedExistingPane = taskPaneManager.TryShowExistingPaneForDisplayRequest(workbook, window);
            bool shouldShowWithRenderPane = !showedExistingPane
                && taskPaneManager.ShouldShowWithRenderPaneForDisplayRequest(workbook, window);

            return Decide(showedExistingPane, shouldShowWithRenderPane);
        }

        internal static PaneDisplayPolicyResult Decide(bool showedExistingPane, bool shouldShowWithRenderPane)
        {
            if (showedExistingPane)
            {
                return PaneDisplayPolicyResult.ShowExisting;
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
    }
}
