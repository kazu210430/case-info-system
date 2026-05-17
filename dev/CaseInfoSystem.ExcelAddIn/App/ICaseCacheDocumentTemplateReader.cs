using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// CASE cache-only の文書テンプレート metadata 参照口。
    /// Base snapshot から CASE cache への promotion と DocProperty 更新を伴う場合があるため、pure read ではない。
    /// </summary>
    internal interface ICaseCacheDocumentTemplateReader
    {
        bool TryEnsurePromotedCaseCacheThenResolve(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result);
    }
}
