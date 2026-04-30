using System;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class TaskPaneSnapshotChunkStorageHelper
    {
        private const int DefaultChunkSize = 240;

        internal static void SaveSnapshot(
            ExcelInteropService excelInteropService,
            Excel.Workbook workbook,
            string countPropName,
            string partPropPrefix,
            string snapshotText,
            int chunkSize = DefaultChunkSize)
        {
            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }

            if (workbook == null
                || string.IsNullOrWhiteSpace(countPropName)
                || string.IsNullOrWhiteSpace(partPropPrefix)
                || chunkSize <= 0)
            {
                return;
            }

            int previousCount = ReadPositiveIntProperty(excelInteropService, workbook, countPropName);
            if (string.IsNullOrEmpty(snapshotText))
            {
                excelInteropService.SetDocumentProperty(workbook, countPropName, "0");
                return;
            }

            int partCount = ((snapshotText.Length - 1) / chunkSize) + 1;
            excelInteropService.SetDocumentProperty(workbook, countPropName, partCount.ToString(CultureInfo.InvariantCulture));

            for (int partIndex = 1; partIndex <= partCount; partIndex++)
            {
                int startIndex = (partIndex - 1) * chunkSize;
                int length = Math.Min(chunkSize, snapshotText.Length - startIndex);
                excelInteropService.SetDocumentProperty(
                    workbook,
                    partPropPrefix + partIndex.ToString("00", CultureInfo.InvariantCulture),
                    snapshotText.Substring(startIndex, length));
            }

            for (int partIndex = partCount + 1; partIndex <= previousCount; partIndex++)
            {
                excelInteropService.SetDocumentProperty(
                    workbook,
                    partPropPrefix + partIndex.ToString("00", CultureInfo.InvariantCulture),
                    string.Empty);
            }
        }

        internal static void ClearSnapshot(
            ExcelInteropService excelInteropService,
            Excel.Workbook workbook,
            string countPropName,
            string partPropPrefix)
        {
            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }

            if (workbook == null || string.IsNullOrWhiteSpace(countPropName) || string.IsNullOrWhiteSpace(partPropPrefix))
            {
                return;
            }

            int previousCount = ReadPositiveIntProperty(excelInteropService, workbook, countPropName);
            excelInteropService.SetDocumentProperty(workbook, countPropName, "0");
            for (int partIndex = 1; partIndex <= previousCount; partIndex++)
            {
                excelInteropService.SetDocumentProperty(
                    workbook,
                    partPropPrefix + partIndex.ToString("00", CultureInfo.InvariantCulture),
                    string.Empty);
            }
        }

        private static int ReadPositiveIntProperty(ExcelInteropService excelInteropService, Excel.Workbook workbook, string propertyName)
        {
            string text = excelInteropService.TryGetDocumentProperty(workbook, propertyName);
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value > 0
                ? value
                : 0;
        }
    }
}
