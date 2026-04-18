using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseListMappingRepository
    {
        private const string SheetName = "CaseList_Mapping";

        private readonly ExcelInteropService _excelInteropService;

        /// <summary>
        internal CaseListMappingRepository(ExcelInteropService excelInteropService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
        }

        /// <summary>
        internal IReadOnlyList<CaseListMappingDefinition> LoadEnabledDefinitions(Excel.Workbook kernelWorkbook)
        {
            var result = new List<CaseListMappingDefinition>();
            Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByName(kernelWorkbook, SheetName);
            if (worksheet == null)
            {
                return result;
            }

            IReadOnlyList<IReadOnlyDictionary<string, string>> rows = _excelInteropService.ReadRecordsFromHeaderRow(worksheet);
            foreach (IReadOnlyDictionary<string, string> row in rows)
            {
                if (!IsEnabled(row))
                {
                    continue;
                }

                string sourceFieldKey = GetValue(row, "SourceFieldKey");
                string targetHeaderName = GetValue(row, "TargetHeaderName");
                if (string.IsNullOrWhiteSpace(sourceFieldKey) || string.IsNullOrWhiteSpace(targetHeaderName))
                {
                    continue;
                }

                result.Add(new CaseListMappingDefinition
                {
                    MappingType = GetValue(row, "MappingType"),
                    SourceFieldKey = sourceFieldKey,
                    TargetHeaderName = targetHeaderName,
                    DataType = GetValue(row, "DataType"),
                    NormalizeRule = GetValue(row, "NormalizeRule")
                });
            }

            return result;
        }

        /// <summary>
        private static bool IsEnabled(IReadOnlyDictionary<string, string> row)
        {
            string enabledValue = GetValue(row, "Enabled").Trim();
            return string.Equals(enabledValue, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(enabledValue, "true", StringComparison.OrdinalIgnoreCase);
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
