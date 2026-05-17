using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// CASE cache を優先しつつ、必要時のみ master catalog へフォールバックする実行側 lookup 参照口。
    /// 先行する CASE cache lookup は promotion-aware で、DocProperty 更新を伴う場合がある。
    /// </summary>
    internal interface IDocumentTemplateLookupReader
    {
        bool TryResolveWithMasterFallback(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result);
    }
}
