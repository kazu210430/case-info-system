using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookEventCoordinator
    {
        private readonly ThisAddIn _addin;
        private readonly AccountingWorkbookLifecycleService _accountingWorkbookLifecycleService;

        internal WorkbookEventCoordinator(ThisAddIn addin, AccountingWorkbookLifecycleService accountingWorkbookLifecycleService)
        {
            _addin = addin;
            _accountingWorkbookLifecycleService = accountingWorkbookLifecycleService;
        }

        internal void OnWindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            _accountingWorkbookLifecycleService?.HandleWindowActivated(
                workbook,
                window,
                AccountingInitialSheetSyncPolicy.WindowActivateEventName);
            _addin.HandleWindowActivateEvent(workbook, window);
        }

        internal void OnWindowActivate(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            _accountingWorkbookLifecycleService?.HandleWindowActivated(
                triggerFacts == null ? null : triggerFacts.Workbook,
                triggerFacts == null ? null : triggerFacts.Window,
                AccountingInitialSheetSyncPolicy.WindowActivateEventName);
            _addin.HandleWindowActivateEvent(triggerFacts);
        }
    }
}
