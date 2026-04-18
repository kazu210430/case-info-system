using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests.Fakes
{
    internal sealed class FakeExcelInteropService : IExcelInteropService
    {
        internal Excel.Workbook ActiveWorkbook { get; set; }

        internal Excel.Window ActiveWindow { get; set; }

        internal string WorkbookFullName { get; set; }

        internal string ActiveSheetCodeName { get; set; }

        internal Excel.Window FirstVisibleWindow { get; set; }

        internal bool ActivateWorkbookResult { get; set; } = true;

        internal bool ActivateWorksheetByCodeNameResult { get; set; } = true;

        public Excel.Workbook GetActiveWorkbook()
        {
            return ActiveWorkbook;
        }

        public Excel.Window GetActiveWindow()
        {
            return ActiveWindow;
        }

        public string GetWorkbookFullName(Excel.Workbook workbook)
        {
            return WorkbookFullName ?? string.Empty;
        }

        public string GetActiveSheetCodeName(Excel.Workbook workbook)
        {
            return ActiveSheetCodeName ?? string.Empty;
        }

        public Excel.Window GetFirstVisibleWindow(Excel.Workbook workbook)
        {
            return FirstVisibleWindow;
        }

        public bool ActivateWorkbook(Excel.Workbook workbook)
        {
            return ActivateWorkbookResult;
        }

        public bool ActivateWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName)
        {
            return ActivateWorksheetByCodeNameResult;
        }
    }

    internal sealed class FakeWorkbookRoleResolver : IWorkbookRoleResolver
    {
        internal WorkbookRole Role { get; set; } = WorkbookRole.Unknown;

        internal string SystemRoot { get; set; } = string.Empty;

        public WorkbookRole Resolve(Excel.Workbook workbook)
        {
            return Role;
        }

        public string ResolveSystemRoot(Excel.Workbook workbook)
        {
            return SystemRoot ?? string.Empty;
        }
    }
}
