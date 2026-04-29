using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// CASE cache だけを読み取り、文書テンプレート metadata を返す read-only API。
    /// </summary>
    internal interface ICaseCacheDocumentTemplateReader
    {
        bool TryResolveFromCaseCache(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result);
    }
}
