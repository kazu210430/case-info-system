using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class ExcelValidationService
    {
        private readonly Logger _logger;

        internal ExcelValidationService(Logger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal IReadOnlyList<string> GetValidationCandidates(Excel.Range targetCell)
        {
            var result = new List<string>();
            if (targetCell == null)
            {
                return result;
            }

            Excel.Validation validation = null;
            Excel.Range sourceRange = null;
            try
            {
                validation = targetCell.Validation;
                string formula1 = Convert.ToString(validation.Formula1) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(formula1))
                {
                    _logger.Debug(nameof(ExcelValidationService), "Validation formula was empty.");
                    return result;
                }

                if (formula1.StartsWith("=", StringComparison.Ordinal))
                {
                    sourceRange = ResolveValidationSourceRange(targetCell, formula1);
                    if (sourceRange == null)
                    {
                        return result;
                    }

                    foreach (Excel.Range cell in sourceRange.Cells)
                    {
                        string text = (Convert.ToString(cell.Value2) ?? string.Empty).Trim();
                        if (text.Length > 0)
                        {
                            result.Add(text);
                        }
                    }

                    _logger.Info(
                        "Validation candidates resolved from range. address="
                        + SafeAddress(targetCell)
                        + ", formula="
                        + formula1
                        + ", count="
                        + result.Count.ToString());
                    return result;
                }

                string[] items = formula1.Split(',');
                for (int index = 0; index < items.Length; index++)
                {
                    string itemText = (items[index] ?? string.Empty).Trim();
                    if (itemText.Length > 0)
                    {
                        result.Add(itemText);
                    }
                }

                _logger.Info(
                    "Validation candidates resolved from inline list. address="
                    + SafeAddress(targetCell)
                    + ", count="
                    + result.Count.ToString());
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error("Validation candidates resolve failed. address=" + SafeAddress(targetCell), ex);
                return result;
            }
            finally
            {
            }
        }

        private static string SafeAddress(Excel.Range range)
        {
            try
            {
                return range == null ? string.Empty : Convert.ToString(range.Address[false, false]) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static Excel.Range ResolveValidationSourceRange(Excel.Range targetCell, string formula1)
        {
            if (targetCell == null)
            {
                return null;
            }

            string expression = (formula1 ?? string.Empty).Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal))
            {
                expression = expression.Substring(1);
            }

            try
            {
                Excel.Worksheet worksheet = targetCell.Worksheet;
                Excel.Workbook workbook = worksheet == null ? null : worksheet.Parent as Excel.Workbook;
                Excel.Names workbookNames = workbook == null ? null : workbook.Names;
                Excel.Names worksheetNames = worksheet == null ? null : worksheet.Names;

                Excel.Range resolvedRange = TryResolveName(workbookNames, expression);
                if (resolvedRange != null)
                {
                    return resolvedRange;
                }

                resolvedRange = TryResolveName(worksheetNames, expression);
                if (resolvedRange != null)
                {
                    return resolvedRange;
                }

                resolvedRange = ResolveSheetQualifiedRange(workbook, expression);
                if (resolvedRange != null)
                {
                    return resolvedRange;
                }

                return worksheet == null ? null : worksheet.Range[expression];
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Range TryResolveName(Excel.Names names, string expression)
        {
            if (names == null || string.IsNullOrWhiteSpace(expression))
            {
                return null;
            }

            try
            {
                Excel.Name name = names.Item(expression);
                return name == null ? null : name.RefersToRange;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Range ResolveSheetQualifiedRange(Excel.Workbook workbook, string expression)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(expression))
            {
                return null;
            }

            int separatorIndex = expression.LastIndexOf('!');
            if (separatorIndex <= 0)
            {
                return null;
            }

            string sheetToken = expression.Substring(0, separatorIndex).Replace("'", string.Empty);
            string addressText = expression.Substring(separatorIndex + 1);

            try
            {
                Excel.Worksheet worksheet = workbook.Worksheets[sheetToken] as Excel.Worksheet;
                return worksheet == null ? null : worksheet.Range[addressText];
            }
            catch
            {
                return null;
            }
        }
    }
}
