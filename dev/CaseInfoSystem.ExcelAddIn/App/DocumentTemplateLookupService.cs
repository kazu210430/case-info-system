using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
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

        public bool TryResolveFromCaseCache(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result)
        {
            return _taskPaneSnapshotCacheService.TryGetDocumentTemplateLookupFromCache(workbook, key, out result);
        }

        public bool TryResolveWithMasterFallback(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result)
        {
            result = null;

            if (TryResolveFromCaseCache(workbook, key, out result))
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
