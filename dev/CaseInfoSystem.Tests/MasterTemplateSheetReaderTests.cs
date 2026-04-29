using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class MasterTemplateSheetReaderTests
    {
        [Fact]
        public void BuildFromValues_ReadsKeyFileCaptionTabAndColors()
        {
            Array values = CreateValues(
                new object[] { "1", "01_申請書.docx", "申請書", "unused", "申請" },
                new object[] { "02", "02_確認書.dotx", "確認書", "unused", string.Empty });

            var fillColors = new Dictionary<int, long>
            {
                [3] = 101L,
                [4] = 202L
            };
            var tabBackColors = new Dictionary<int, long>
            {
                [3] = 303L,
                [4] = 404L
            };

            MasterTemplateSheetData result = MasterTemplateSheetReader.BuildFromValues(
                4,
                values,
                rowIndex => fillColors[rowIndex],
                rowIndex => tabBackColors[rowIndex]);

            Assert.Equal(4, result.LastRow);
            Assert.Collection(
                result.Rows,
                first =>
                {
                    Assert.Equal(3, first.RowIndex);
                    Assert.Equal("01", first.Key);
                    Assert.Equal("01_申請書.docx", first.TemplateFileName);
                    Assert.Equal("申請書", first.Caption);
                    Assert.Equal("申請", first.TabName);
                    Assert.Equal(101L, first.FillColor);
                    Assert.Equal(303L, first.TabBackColor);
                },
                second =>
                {
                    Assert.Equal(4, second.RowIndex);
                    Assert.Equal("02", second.Key);
                    Assert.Equal("02_確認書.dotx", second.TemplateFileName);
                    Assert.Equal("確認書", second.Caption);
                    Assert.Equal(string.Empty, second.TabName);
                    Assert.Equal(202L, second.FillColor);
                    Assert.Equal(404L, second.TabBackColor);
                });
        }

        [Fact]
        public void BuildFromValues_PreservesEmptyAndIncompleteRowsForCallers()
        {
            Array values = CreateValues(
                new object[] { "", "", "", "", "" },
                new object[] { "3", "", "見出しのみ", "unused", "補助" });

            MasterTemplateSheetData result = MasterTemplateSheetReader.BuildFromValues(
                4,
                values,
                rowIndex => 0L,
                rowIndex => 0L);

            Assert.Equal(2, result.Rows.Count);

            MasterTemplateSheetRowData emptyRow = result.Rows[0];
            Assert.Equal(3, emptyRow.RowIndex);
            Assert.Equal(string.Empty, emptyRow.Key);
            Assert.Equal(string.Empty, emptyRow.TemplateFileName);
            Assert.Equal(string.Empty, emptyRow.Caption);
            Assert.Equal(string.Empty, emptyRow.TabName);

            MasterTemplateSheetRowData incompleteRow = result.Rows[1];
            Assert.Equal(4, incompleteRow.RowIndex);
            Assert.Equal("03", incompleteRow.Key);
            Assert.Equal(string.Empty, incompleteRow.TemplateFileName);
            Assert.Equal("見出しのみ", incompleteRow.Caption);
            Assert.Equal("補助", incompleteRow.TabName);
        }

        private static Array CreateValues(params object[][] rows)
        {
            Array values = Array.CreateInstance(typeof(object), new[] { rows.Length, 5 }, new[] { 1, 1 });
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                object[] row = rows[rowIndex] ?? Array.Empty<object>();
                for (int columnIndex = 0; columnIndex < 5; columnIndex++)
                {
                    object value = columnIndex < row.Length ? row[columnIndex] : null;
                    values.SetValue(value, rowIndex + 1, columnIndex + 1);
                }
            }

            return values;
        }
    }
}
