using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal interface IWorkbookRoleResolver
    {
        WorkbookRole Resolve(Excel.Workbook workbook);

        string ResolveSystemRoot(Excel.Workbook workbook);
    }
}
