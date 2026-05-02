using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal interface ICasePaneHostBridge
    {
        void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason);

        void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason);

        bool ShouldIgnoreWorkbookActivateDuringCaseProtection(Excel.Workbook workbook);

        bool ShouldIgnoreTaskPaneRefreshDuringCaseProtection(string reason, Excel.Workbook workbook, Excel.Window window);

        bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window);

        void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason);
    }

    internal sealed class ThisAddInCasePaneHostBridge : ICasePaneHostBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInCasePaneHostBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason)
        {
            _addIn.SuppressUpcomingCasePaneActivationRefresh(workbookFullName, reason);
        }

        public void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            _addIn.ShowWorkbookTaskPaneWhenReady(workbook, reason);
        }

        public bool ShouldIgnoreWorkbookActivateDuringCaseProtection(Excel.Workbook workbook)
        {
            return _addIn.ShouldIgnoreWorkbookActivateDuringCaseProtection(workbook);
        }

        public bool ShouldIgnoreTaskPaneRefreshDuringCaseProtection(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _addIn.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window);
        }

        public bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            return _addIn.HasVisibleCasePaneForWorkbookWindow(workbook, window);
        }

        public void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            _addIn.BeginCaseWorkbookActivateProtection(workbook, window, reason);
        }
    }
}
