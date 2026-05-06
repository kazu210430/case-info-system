using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookCaseTaskPaneRefreshCommandService
    {
        private const string ProductTitle = "案件情報System";

        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Func<string, Excel.Workbook, Excel.Window, bool> _isTaskPaneRefreshSucceeded;

        internal WorkbookCaseTaskPaneRefreshCommandService(
            WorkbookRoleResolver workbookRoleResolver,
            ExcelInteropService excelInteropService,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Func<string, Excel.Workbook, Excel.Window, bool> isTaskPaneRefreshSucceeded)
        {
            _workbookRoleResolver = workbookRoleResolver;
            _excelInteropService = excelInteropService;
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow;
            _isTaskPaneRefreshSucceeded = isTaskPaneRefreshSucceeded;
        }

        internal void Refresh(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                UserErrorService.ShowOkNotification("対象ブックを取得できませんでした。", ProductTitle, MessageBoxIcon.Information);
                return;
            }

            if (_workbookRoleResolver == null || _excelInteropService == null || _resolveWorkbookPaneWindow == null || _isTaskPaneRefreshSucceeded == null)
            {
                UserErrorService.ShowOkNotification("Pane 更新サービスを利用できません。", ProductTitle, MessageBoxIcon.Warning);
                return;
            }

            WorkbookRole role = _workbookRoleResolver.Resolve(workbook);
            if (role != WorkbookRole.Case)
            {
                UserErrorService.ShowOkNotification("CASE ブックで実行してください。", ProductTitle, MessageBoxIcon.Information);
                return;
            }

            Excel.Window window = _resolveWorkbookPaneWindow(workbook, "RibbonCasePaneRefresh", true);
            bool refreshed = _isTaskPaneRefreshSucceeded("RibbonCasePaneRefresh", workbook, window);
            if (!refreshed)
            {
                UserErrorService.ShowOkNotification("文書ボタンパネルを更新できませんでした。", ProductTitle, MessageBoxIcon.Warning);
            }
        }
    }
}
