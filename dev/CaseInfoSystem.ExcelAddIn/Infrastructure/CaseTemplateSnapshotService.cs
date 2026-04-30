using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseTemplateSnapshotService
    {
        private const string TaskPaneMasterVersionProp = "TASKPANE_MASTER_VERSION";
        private const string TaskPaneCacheCountProp = "TASKPANE_SNAPSHOT_CACHE_COUNT";
        private const string TaskPaneCachePartPropPrefix = "TASKPANE_SNAPSHOT_CACHE_";
        private const string TaskPaneBaseCacheCountProp = "TASKPANE_BASE_SNAPSHOT_COUNT";
        private const string TaskPaneBaseCachePartPropPrefix = "TASKPANE_BASE_SNAPSHOT_";
        private const string TaskPaneBaseMasterVersionProp = "TASKPANE_BASE_MASTER_VERSION";

        private readonly ExcelInteropService _excelInteropService;

        internal CaseTemplateSnapshotService(ExcelInteropService excelInteropService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
        }

        internal void SyncMasterVersionFromKernel(Excel.Workbook kernelWorkbook, Excel.Workbook caseWorkbook)
        {
            string masterVersion = _excelInteropService.TryGetDocumentProperty(kernelWorkbook, TaskPaneMasterVersionProp);
            if (!string.IsNullOrWhiteSpace(masterVersion))
            {
                _excelInteropService.SetDocumentProperty(caseWorkbook, TaskPaneMasterVersionProp, masterVersion);
            }
        }

        internal void PromoteEmbeddedSnapshotToCaseCache(Excel.Workbook caseWorkbook)
        {
            if (caseWorkbook == null)
            {
                throw new ArgumentNullException(nameof(caseWorkbook));
            }

            int embeddedCount = ReadPositiveIntProperty(caseWorkbook, TaskPaneBaseCacheCountProp);
            if (embeddedCount <= 0)
            {
                return;
            }

            string embeddedSnapshot = LoadSnapshot(caseWorkbook, TaskPaneBaseCacheCountProp, TaskPaneBaseCachePartPropPrefix);
            if (!TaskPaneSnapshotFormat.IsCompatible(embeddedSnapshot))
            {
                ClearSnapshot(caseWorkbook, TaskPaneBaseCacheCountProp, TaskPaneBaseCachePartPropPrefix);
                ClearSnapshot(caseWorkbook, TaskPaneCacheCountProp, TaskPaneCachePartPropPrefix);
                return;
            }

            int previousCaseCacheCount = ReadPositiveIntProperty(caseWorkbook, TaskPaneCacheCountProp);
            _excelInteropService.SetDocumentProperty(caseWorkbook, TaskPaneCacheCountProp, embeddedCount.ToString());
            for (int partIndex = 1; partIndex <= embeddedCount; partIndex++)
            {
                string sourceProp = TaskPaneBaseCachePartPropPrefix + partIndex.ToString("00");
                string targetProp = TaskPaneCachePartPropPrefix + partIndex.ToString("00");
                string partText = _excelInteropService.TryGetDocumentProperty(caseWorkbook, sourceProp);
                _excelInteropService.SetDocumentProperty(caseWorkbook, targetProp, partText);
            }

            for (int partIndex = embeddedCount + 1; partIndex <= previousCaseCacheCount; partIndex++)
            {
                _excelInteropService.SetDocumentProperty(caseWorkbook, TaskPaneCachePartPropPrefix + partIndex.ToString("00"), string.Empty);
            }

            int embeddedMasterVersion = ReadPositiveIntProperty(caseWorkbook, TaskPaneBaseMasterVersionProp);
            if (embeddedMasterVersion > 0)
            {
                _excelInteropService.SetDocumentProperty(caseWorkbook, TaskPaneMasterVersionProp, embeddedMasterVersion.ToString());
            }
        }

        private int ReadPositiveIntProperty(Excel.Workbook workbook, string propertyName)
        {
            string text = _excelInteropService.TryGetDocumentProperty(workbook, propertyName);
            return int.TryParse(text, out int value) && value > 0 ? value : 0;
        }

        /// <summary>
        private string LoadSnapshot(Excel.Workbook workbook, string countPropName, string partPropPrefix)
        {
            return TaskPaneSnapshotChunkReadHelper.LoadSnapshot(_excelInteropService, workbook, countPropName, partPropPrefix);
        }

        /// <summary>
        private void ClearSnapshot(Excel.Workbook workbook, string countPropName, string partPropPrefix)
        {
            int previousCount = ReadPositiveIntProperty(workbook, countPropName);
            _excelInteropService.SetDocumentProperty(workbook, countPropName, "0");
            for (int partIndex = 1; partIndex <= previousCount; partIndex++)
            {
                _excelInteropService.SetDocumentProperty(workbook, partPropPrefix + partIndex.ToString("00"), string.Empty);
            }
        }
    }
}
