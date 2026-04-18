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

        internal void OnWorkbookOpen(Excel.Workbook workbook)
        {
            _addin.HandleWorkbookOpenEvent(workbook);
        }

        internal void OnWorkbookActivate(Excel.Workbook workbook)
        {
            _addin.HandleWorkbookActivateEvent(workbook);
        }

        internal void OnWindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            _addin.HandleWindowActivateEvent(workbook, window);
        }
    }
}
