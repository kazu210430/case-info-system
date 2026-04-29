using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// CASE cache を優先しつつ、必要時のみ master catalog へフォールバックする read-only 参照口。
    /// </summary>
    internal interface IDocumentTemplateLookupReader
    {
        bool TryResolveWithMasterFallback(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result);
    }
}
