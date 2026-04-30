using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    /// Master catalog から文書テンプレート metadata を key 単位で読み取る read-only API。
    /// </summary>
    internal interface IMasterTemplateCatalogReader
    {
        bool TryGetTemplateByKey(Excel.Workbook caseWorkbook, string key, out MasterTemplateRecord record);
    }
}
