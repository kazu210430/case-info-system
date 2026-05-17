using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// prompt 用の CASE cache-only lookup と、文書実行用の master fallback lookup を分ける。
    /// CASE cache leg は Base snapshot promotion を含み得るため pure read ではない。
    /// </summary>
    internal sealed class DocumentTemplateLookupService : ICaseCacheDocumentTemplateReader, IDocumentTemplateLookupReader
    {
        private readonly TaskPaneSnapshotCacheService _taskPaneSnapshotCacheService;
        private readonly IMasterTemplateCatalogReader _masterTemplateCatalogReader;

        internal DocumentTemplateLookupService(
            TaskPaneSnapshotCacheService taskPaneSnapshotCacheService,
            IMasterTemplateCatalogReader masterTemplateCatalogReader)
        {
            _taskPaneSnapshotCacheService = taskPaneSnapshotCacheService ?? throw new ArgumentNullException(nameof(taskPaneSnapshotCacheService));
            _masterTemplateCatalogReader = masterTemplateCatalogReader ?? throw new ArgumentNullException(nameof(masterTemplateCatalogReader));
        }

        public bool TryEnsurePromotedCaseCacheThenResolve(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result)
        {
            return _taskPaneSnapshotCacheService.TryEnsurePromotedCaseCacheThenGetDocumentTemplateLookup(workbook, key, out result);
        }

        public bool TryResolveWithMasterFallback(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result)
        {
            result = null;

            if (TryEnsurePromotedCaseCacheThenResolve(workbook, key, out result))
            {
                return true;
            }

            if (workbook == null)
            {
                return false;
            }

            if (!_masterTemplateCatalogReader.TryGetTemplateByKey(workbook, key, out MasterTemplateRecord masterRecord))
            {
                return false;
            }

            result = new DocumentTemplateLookupResult
            {
                Key = masterRecord.Key ?? string.Empty,
                DocumentName = masterRecord.DocumentName ?? string.Empty,
                TemplateFileName = masterRecord.TemplateFileName ?? string.Empty,
                ResolutionSource = DocumentTemplateResolutionSource.MasterCatalog
            };
            return result.TemplateFileName.Length > 0;
        }
    }
}
