using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class UserDataBaseMappingRepository
    {
        private const string SheetName = "UserData_BaseMapping";

        private readonly ExcelInteropService _excelInteropService;

        /// <summary>
        internal UserDataBaseMappingRepository(ExcelInteropService excelInteropService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
        }

        /// <summary>
        internal IReadOnlyList<UserDataBaseMappingDefinition> LoadEnabledDefinitions(Excel.Workbook kernelWorkbook)
        {
            var result = new List<UserDataBaseMappingDefinition>();
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
                string targetFieldKey = GetValue(row, "TargetFieldKey");
                if (string.IsNullOrWhiteSpace(sourceFieldKey) || string.IsNullOrWhiteSpace(targetFieldKey))
                {
                    continue;
                }

                result.Add(new UserDataBaseMappingDefinition
                {
                    SourceFieldKey = sourceFieldKey,
                    TargetFieldKey = targetFieldKey
                });
            }

            return result;
        }

        /// <summary>
        private static bool IsEnabled(IReadOnlyDictionary<string, string> row)
        {
            string enabled = GetValue(row, "Enabled").Trim();
            return string.Equals(enabled, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(enabled, "true", StringComparison.OrdinalIgnoreCase);
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
