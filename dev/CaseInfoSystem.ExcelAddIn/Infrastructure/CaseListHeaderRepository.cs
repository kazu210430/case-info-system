using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseListHeaderRepository
    {
        private const string SheetName = "CaseList_Headers";

        private readonly ExcelInteropService _excelInteropService;

        /// <summary>
        internal CaseListHeaderRepository(ExcelInteropService excelInteropService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
        }

        /// <summary>
        internal IReadOnlyList<CaseListHeaderDefinition> LoadDefinitions(Excel.Workbook kernelWorkbook)
        {
            var result = new List<CaseListHeaderDefinition>();
            Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByName(kernelWorkbook, SheetName);
            if (worksheet == null)
            {
                return result;
            }

            IReadOnlyList<IReadOnlyDictionary<string, string>> rows = _excelInteropService.ReadRecordsFromHeaderRow(worksheet);
            foreach (IReadOnlyDictionary<string, string> row in rows)
            {
                string headerName = GetValue(row, "Header");
                if (string.IsNullOrWhiteSpace(headerName))
                {
                    continue;
                }

                result.Add(new CaseListHeaderDefinition
                {
                    CellAddress = GetValue(row, "Cell"),
                    HeaderName = headerName
                });
            }

            return result;
        }

        /// <summary>
        private static string GetValue(IReadOnlyDictionary<string, string> row, string columnName)
        {
            if (row == null || string.IsNullOrWhiteSpace(columnName))
            {
                return string.Empty;
            }

            string value;
            return row.TryGetValue(columnName, out value) ? value ?? string.Empty : string.Empty;
        }
    }
}
