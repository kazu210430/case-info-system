using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal interface IExcelInteropService
    {
        Excel.Workbook GetActiveWorkbook();

        Excel.Window GetActiveWindow();

        string GetWorkbookFullName(Excel.Workbook workbook);

        string GetActiveSheetCodeName(Excel.Workbook workbook);

        Excel.Window GetFirstVisibleWindow(Excel.Workbook workbook);

        bool ActivateWorkbook(Excel.Workbook workbook);

        bool ActivateWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName);
    }
}
