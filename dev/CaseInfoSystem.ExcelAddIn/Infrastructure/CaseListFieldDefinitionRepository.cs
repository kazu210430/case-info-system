using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseListFieldDefinitionRepository
    {
        private const string SheetName = "CaseList_FieldInventory";

        private readonly ExcelInteropService _excelInteropService;

        /// <summary>
        internal CaseListFieldDefinitionRepository(ExcelInteropService excelInteropService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
        }

        /// <summary>
        internal IReadOnlyDictionary<string, CaseListFieldDefinition> LoadDefinitions(Excel.Workbook kernelWorkbook)
        {
            var result = new Dictionary<string, CaseListFieldDefinition>(StringComparer.OrdinalIgnoreCase);
            foreach (CaseListFieldDefinition definition in LoadDefinitionList(kernelWorkbook))
            {
                if (definition == null || string.IsNullOrWhiteSpace(definition.FieldKey))
                {
                    continue;
                }

                if (!result.ContainsKey(definition.FieldKey))
                {
                    result[definition.FieldKey] = definition;
                }
            }

            return result;
        }

        /// <summary>
        internal IReadOnlyList<CaseListFieldDefinition> LoadDefinitionList(Excel.Workbook kernelWorkbook)
        {
            var result = new List<CaseListFieldDefinition>();
            Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByName(kernelWorkbook, SheetName);
            if (worksheet == null)
            {
                return result;
            }

            IReadOnlyList<IReadOnlyDictionary<string, string>> rows = _excelInteropService.ReadRecordsFromHeaderRow(worksheet);
            foreach (IReadOnlyDictionary<string, string> row in rows)
            {
                string fieldKey = GetValue(row, "ProposedFieldKey");
                if (string.IsNullOrWhiteSpace(fieldKey))
                {
                    continue;
                }

                result.Add(new CaseListFieldDefinition
                {
                    FieldKey = fieldKey,
                    Label = GetValue(row, "Label"),
                    SourceCellAddress = GetValue(row, "SourceCell"),
                    SourceNamedRange = GetValue(row, "ProposedNamedRange"),
                    DataType = GetValue(row, "DataType"),
                    NormalizeRule = GetValue(row, "NormalizeRule")
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
