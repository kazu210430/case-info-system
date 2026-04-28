using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.UI;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class TaskPaneSnapshotCacheService
    {
        private const string TaskPaneCacheCountProp = "TASKPANE_SNAPSHOT_CACHE_COUNT";
        private const string TaskPaneCachePartPropPrefix = "TASKPANE_SNAPSHOT_CACHE_";
        private const string TaskPaneBaseCacheCountProp = "TASKPANE_BASE_SNAPSHOT_COUNT";
        private const string TaskPaneBaseCachePartPropPrefix = "TASKPANE_BASE_SNAPSHOT_";
        private const string TaskPaneBaseMasterVersionProp = "TASKPANE_BASE_MASTER_VERSION";
        private const string TaskPaneMasterVersionProp = "TASKPANE_MASTER_VERSION";

        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        internal TaskPaneSnapshotCacheService(ExcelInteropService excelInteropService, Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal bool PromoteBaseSnapshotToCaseCacheIfNeeded(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            string currentSnapshot = LoadTaskPaneSnapshotCache(workbook);
            string embeddedSnapshot = LoadTaskPaneBaseSnapshotCache(workbook);
            if (currentSnapshot.Length > 0 && !TaskPaneSnapshotFormat.IsCompatible(currentSnapshot))
            {
                ClearSnapshotParts(workbook, TaskPaneCacheCountProp, TaskPaneCachePartPropPrefix);
                currentSnapshot = string.Empty;
            }

            if (embeddedSnapshot.Length > 0 && !TaskPaneSnapshotFormat.IsCompatible(embeddedSnapshot))
            {
                ClearSnapshotParts(workbook, TaskPaneBaseCacheCountProp, TaskPaneBaseCachePartPropPrefix);
                embeddedSnapshot = string.Empty;
            }

            if (embeddedSnapshot.Length == 0)
            {
                return false;
            }

            int caseMasterVersion = ReadPositiveIntProperty(workbook, TaskPaneMasterVersionProp);
            int embeddedMasterVersion = ReadPositiveIntProperty(workbook, TaskPaneBaseMasterVersionProp);

            bool shouldPromote = currentSnapshot.Length == 0;
            if (!shouldPromote && embeddedMasterVersion > 0)
            {
                shouldPromote = caseMasterVersion <= 0 || embeddedMasterVersion > caseMasterVersion;
            }

            if (!shouldPromote)
            {
                return false;
            }

            SaveTaskPaneSnapshotCache(workbook, embeddedSnapshot);
            if (embeddedMasterVersion > 0)
            {
                _excelInteropService.SetDocumentProperty(workbook, TaskPaneMasterVersionProp, embeddedMasterVersion.ToString());
            }

            _logger.Info(
                "TaskPaneSnapshotCacheService promoted Base snapshot to CASE cache. "
                + "caseMasterVersion="
                + caseMasterVersion.ToString()
                + ", embeddedMasterVersion="
                + embeddedMasterVersion.ToString()
                + ", hadExistingCaseCache="
                + (currentSnapshot.Length > 0).ToString());
            return true;
        }

        internal bool TryGetDocInfoFromCache(Excel.Workbook workbook, string key, out string templateFileName, out string documentName)
        {
            templateFileName = string.Empty;
            documentName = string.Empty;

            if (workbook == null)
            {
                return false;
            }

            PromoteBaseSnapshotToCaseCacheIfNeeded(workbook);
            string normalizedKey = NormalizeDocButtonKey(key);
            if (normalizedKey.Length == 0)
            {
                return false;
            }

            string snapshotText = LoadTaskPaneSnapshotCache(workbook);
            if (snapshotText.Length == 0)
            {
                return false;
            }

            if (!TaskPaneSnapshotFormat.IsCompatible(snapshotText))
            {
                ClearSnapshotParts(workbook, TaskPaneCacheCountProp, TaskPaneCachePartPropPrefix);
                return false;
            }

            TaskPaneSnapshot snapshot = TaskPaneSnapshotParser.Parse(snapshotText);
            if (snapshot == null || snapshot.DocButtons == null)
            {
                return false;
            }

            foreach (TaskPaneDocDefinition definition in snapshot.DocButtons)
            {
                if (definition == null || !string.Equals(NormalizeDocButtonKey(definition.Key), normalizedKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                documentName = definition.Caption ?? string.Empty;
                templateFileName = definition.TemplateFileName ?? string.Empty;
                return templateFileName.Length > 0;
            }

            return false;
        }

        internal void ClearCaseSnapshotCacheChunks(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            object properties = null;
            try
            {
                properties = workbook.CustomDocumentProperties;
                if (!(properties is DocumentProperties documentProperties))
                {
                    return;
                }

                var propertyNamesToDelete = new List<string>();
                foreach (DocumentProperty item in documentProperties)
                {
                    try
                    {
                        string propertyName = Convert.ToString(item.Name) ?? string.Empty;
                        if (IsCaseSnapshotChunkPropertyName(propertyName))
                        {
                            propertyNamesToDelete.Add(propertyName);
                        }
                    }
                    finally
                    {
                        ReleaseComObject(item);
                    }
                }

                if (propertyNamesToDelete.Count == 0)
                {
                    return;
                }

                dynamic dynamicProperties = properties;
                for (int index = 0; index < propertyNamesToDelete.Count; index++)
                {
                    object property = null;
                    try
                    {
                        string propertyName = propertyNamesToDelete[index];
                        property = dynamicProperties[propertyName];
                        dynamic dynamicProperty = property;
                        dynamicProperty.Delete();
                    }
                    finally
                    {
                        ReleaseComObject(property);
                    }
                }

                _logger.Info("TaskPaneSnapshotCacheService cleared CASE cache chunks. removedCount=" + propertyNamesToDelete.Count.ToString());
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        private string LoadTaskPaneSnapshotCache(Excel.Workbook workbook)
        {
            return LoadSnapshotParts(workbook, TaskPaneCacheCountProp, TaskPaneCachePartPropPrefix);
        }

        private string LoadTaskPaneBaseSnapshotCache(Excel.Workbook workbook)
        {
            return LoadSnapshotParts(workbook, TaskPaneBaseCacheCountProp, TaskPaneBaseCachePartPropPrefix);
        }

        private string LoadSnapshotParts(Excel.Workbook workbook, string countPropName, string partPropPrefix)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            string countText = _excelInteropService.TryGetDocumentProperty(workbook, countPropName);
            if (!int.TryParse(countText, out int partCount) || partCount <= 0)
            {
                return string.Empty;
            }

            var builder = new System.Text.StringBuilder();
            for (int partIndex = 1; partIndex <= partCount; partIndex++)
            {
                builder.Append(_excelInteropService.TryGetDocumentProperty(workbook, partPropPrefix + partIndex.ToString("00")));
            }

            return builder.ToString();
        }

        private void SaveTaskPaneSnapshotCache(Excel.Workbook workbook, string snapshotText)
        {
            if (workbook == null)
            {
                return;
            }

            string previousCountText = _excelInteropService.TryGetDocumentProperty(workbook, TaskPaneCacheCountProp);
            int previousCount = int.TryParse(previousCountText, out int parsedCount) ? parsedCount : 0;

            if (string.IsNullOrEmpty(snapshotText))
            {
                _excelInteropService.SetDocumentProperty(workbook, TaskPaneCacheCountProp, "0");
                return;
            }

            const int taskPaneCacheChunkSize = 240;
            int partCount = ((snapshotText.Length - 1) / taskPaneCacheChunkSize) + 1;
            _excelInteropService.SetDocumentProperty(workbook, TaskPaneCacheCountProp, partCount.ToString());

            for (int partIndex = 1; partIndex <= partCount; partIndex++)
            {
                int startIndex = (partIndex - 1) * taskPaneCacheChunkSize;
                int length = Math.Min(taskPaneCacheChunkSize, snapshotText.Length - startIndex);
                _excelInteropService.SetDocumentProperty(
                    workbook,
                    TaskPaneCachePartPropPrefix + partIndex.ToString("00"),
                    snapshotText.Substring(startIndex, length));
            }

            for (int partIndex = partCount + 1; partIndex <= previousCount; partIndex++)
            {
                _excelInteropService.SetDocumentProperty(workbook, TaskPaneCachePartPropPrefix + partIndex.ToString("00"), string.Empty);
            }
        }

        private static string NormalizeDocButtonKey(string key)
        {
            string trimmed = (key ?? string.Empty).Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            return long.TryParse(trimmed, out long numericKey)
                ? numericKey.ToString("00")
                : trimmed;
        }

        private static bool IsCaseSnapshotChunkPropertyName(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName)
                || !propertyName.StartsWith(TaskPaneCachePartPropPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string suffix = propertyName.Substring(TaskPaneCachePartPropPrefix.Length);
            if (suffix.Length == 0)
            {
                return false;
            }

            for (int index = 0; index < suffix.Length; index++)
            {
                if (!char.IsDigit(suffix[index]))
                {
                    return false;
                }
            }

            return true;
        }

        private int ReadPositiveIntProperty(Excel.Workbook workbook, string propertyName)
        {
            string text = _excelInteropService.TryGetDocumentProperty(workbook, propertyName);
            return int.TryParse(text, out int value) && value > 0 ? value : 0;
        }

        /// <summary>
        private void ClearSnapshotParts(Excel.Workbook workbook, string countPropName, string partPropPrefix)
        {
            if (workbook == null)
            {
                return;
            }

            int previousCount = ReadPositiveIntProperty(workbook, countPropName);
            _excelInteropService.SetDocumentProperty(workbook, countPropName, "0");
            for (int partIndex = 1; partIndex <= previousCount; partIndex++)
            {
                _excelInteropService.SetDocumentProperty(workbook, partPropPrefix + partIndex.ToString("00"), string.Empty);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            Marshal.ReleaseComObject(comObject);
        }
    }
}
