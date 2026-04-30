using System;
using System.Globalization;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class TaskPaneSnapshotChunkReadHelper
    {
        internal static string LoadSnapshot(ExcelInteropService excelInteropService, Excel.Workbook workbook, string countPropName, string partPropPrefix)
        {
            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }

            if (workbook == null || string.IsNullOrWhiteSpace(countPropName) || string.IsNullOrWhiteSpace(partPropPrefix))
            {
                return string.Empty;
            }

            string countText = excelInteropService.TryGetDocumentProperty(workbook, countPropName);
            if (!int.TryParse(countText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int partCount) || partCount <= 0)
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            for (int partIndex = 1; partIndex <= partCount; partIndex++)
            {
                builder.Append(excelInteropService.TryGetDocumentProperty(workbook, partPropPrefix + partIndex.ToString("00", CultureInfo.InvariantCulture)));
            }

            return builder.ToString();
        }
    }
}
