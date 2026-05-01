using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class MasterTemplateSheetData
    {
        internal int LastRow { get; }

        internal IReadOnlyList<MasterTemplateSheetRowData> Rows { get; }

        internal MasterTemplateSheetData(int lastRow, List<MasterTemplateSheetRowData> rows)
        {
            LastRow = lastRow;
            Rows = rows ?? new List<MasterTemplateSheetRowData>();
        }
    }

    internal sealed class MasterTemplateSheetRowData
    {
        internal int RowIndex { get; }

        internal string Key { get; }

        internal string TemplateFileName { get; }

        internal string Caption { get; }

        internal string TabName { get; }

        internal long FillColor { get; }

        internal long TabBackColor { get; }

        internal MasterTemplateSheetRowData(int rowIndex, string key, string templateFileName, string caption, string tabName, long fillColor, long tabBackColor)
        {
            RowIndex = rowIndex;
            Key = key ?? string.Empty;
            TemplateFileName = templateFileName ?? string.Empty;
            Caption = caption ?? string.Empty;
            TabName = tabName ?? string.Empty;
            FillColor = fillColor;
            TabBackColor = tabBackColor;
        }
    }

    internal static class MasterTemplateSheetReader
    {
        private const int FirstDataRow = 3;
        private const int XlUp = -4162;

        internal static MasterTemplateSheetData Read(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                return new MasterTemplateSheetData(0, new List<MasterTemplateSheetRowData>());
            }

            Excel.Range lastCell = null;
            Excel.Range lastUsedCell = null;
            Excel.Range valuesRange = null;
            try
            {
                lastCell = worksheet.Cells[((dynamic)worksheet).Rows.Count, "A"] as Excel.Range;
                if (lastCell == null)
                {
                    return new MasterTemplateSheetData(0, new List<MasterTemplateSheetRowData>());
                }

                lastUsedCell = ((dynamic)lastCell).End[(object)XlUp] as Excel.Range;
                int lastRow = lastUsedCell == null ? 0 : Convert.ToInt32(((dynamic)lastUsedCell).Row);
                if (lastRow < FirstDataRow)
                {
                    return new MasterTemplateSheetData(lastRow, new List<MasterTemplateSheetRowData>());
                }

                valuesRange = ((dynamic)worksheet).Range[(object)("A" + FirstDataRow.ToString()), (object)("E" + lastRow.ToString())] as Excel.Range;
                return BuildFromValues(
                    lastRow,
                    valuesRange == null ? null : valuesRange.Value2 as Array,
                    rowIndex => GetCellInteriorColor(worksheet, rowIndex, "D"),
                    rowIndex => GetCellInteriorColor(worksheet, rowIndex, "F"));
            }
            finally
            {
                ReleaseComObject(valuesRange);
                ReleaseComObject(lastUsedCell);
                ReleaseComObject(lastCell);
            }
        }

        internal static MasterTemplateSheetData BuildFromValues(int lastRow, Array values, Func<int, long> fillColorProvider, Func<int, long> tabBackColorProvider)
        {
            if (lastRow < FirstDataRow || values == null || values.Rank != 2)
            {
                return new MasterTemplateSheetData(lastRow, new List<MasterTemplateSheetRowData>());
            }

            int rowLowerBound = values.GetLowerBound(0);
            int rowUpperBound = values.GetUpperBound(0);
            int columnLowerBound = values.GetLowerBound(1);
            var rows = new List<MasterTemplateSheetRowData>(Math.Max(0, rowUpperBound - rowLowerBound + 1));
            for (int row = rowLowerBound; row <= rowUpperBound; row++)
            {
                int rowIndex = FirstDataRow + (row - rowLowerBound);
                string key = NormalizeDocKey(Convert.ToString(values.GetValue(row, columnLowerBound)));
                string templateFileName = (Convert.ToString(values.GetValue(row, columnLowerBound + 1)) ?? string.Empty).Trim();
                string caption = (Convert.ToString(values.GetValue(row, columnLowerBound + 2)) ?? string.Empty).Trim();
                string tabName = (Convert.ToString(values.GetValue(row, columnLowerBound + 4)) ?? string.Empty).Trim();
                long fillColor = fillColorProvider == null ? 0L : fillColorProvider(rowIndex);
                long tabBackColor = tabBackColorProvider == null ? 0L : tabBackColorProvider(rowIndex);
                rows.Add(new MasterTemplateSheetRowData(rowIndex, key, templateFileName, caption, tabName, fillColor, tabBackColor));
            }

            return new MasterTemplateSheetData(lastRow, rows);
        }

        private static string NormalizeDocKey(string key)
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

        private static long GetCellInteriorColor(Excel.Worksheet worksheet, int rowIndex, string columnName)
        {
            Excel.Range cell = null;
            object interior = null;
            try
            {
                cell = worksheet.Cells[rowIndex, columnName] as Excel.Range;
                interior = cell == null ? null : ((dynamic)cell).Interior;
                return Convert.ToInt64(interior == null ? 0 : ((dynamic)interior).Color);
            }
            finally
            {
                ReleaseComObject(interior);
                ReleaseComObject(cell);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            // Master sheet 読み取りで所有した COM 参照は完全解放の方針を維持する。
            ComObjectReleaseService.FinalRelease(comObject);
        }
    }
}
