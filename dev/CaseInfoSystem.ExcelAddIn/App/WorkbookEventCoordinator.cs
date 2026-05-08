using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookEventCoordinator
    {
        private readonly ThisAddIn _addin;

        internal WorkbookEventCoordinator(ThisAddIn addin)
        {
            _addin = addin;
        }

        internal void OnWindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            _addin.HandleWindowActivateEvent(workbook, window);
        }

        internal void OnWindowActivate(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            _addin.HandleWindowActivateEvent(triggerFacts);
        }
    }
}
