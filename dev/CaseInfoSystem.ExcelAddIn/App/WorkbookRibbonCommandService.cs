using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// Class: coordinates workbook-scoped ribbon commands.
    /// Responsibility: safely show CustomDocumentProperties and update SYSTEM_ROOT.
    /// </summary>
    internal sealed class WorkbookRibbonCommandService
    {
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string WordTemplateDirectoryPropertyName = "WORD_TEMPLATE_DIR";
        private const string TemplateFolderName = "\u96DB\u5F62";
        private const string ProductTitle = "\u6848\u4EF6\u60C5\u5831System";
        private const string SourceSheetCodeName = "shSample";
        private const string TargetSheetCodeName = "shHOME";
        private const int TargetColumnIndex = 2;

        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;

        /// <summary>
        /// Method: initializes the ribbon command service.
        /// Args: excelInteropService - Excel/DocProp service, pathCompatibilityService - path service, logger - technical logger.
        /// Returns: none.
        /// Side effects: none.
        /// </summary>
        internal WorkbookRibbonCommandService(
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Method: shows CustomDocumentProperties for the target workbook.
        /// Args: workbook - target workbook.
        /// Returns: none.
        /// Side effects: shows a dialog or a user message.
        /// </summary>
        internal void ShowCustomDocumentProperties(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                MessageBox.Show("\u5BFE\u8C61\u30D6\u30C3\u30AF\u3092\u53D6\u5F97\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                IReadOnlyList<KeyValuePair<string, string>> properties = _excelInteropService.GetCustomDocumentProperties(workbook);
                string workbookName = _excelInteropService.GetWorkbookName(workbook);
                string content = BuildCustomDocumentPropertyText(properties);

                using (var form = new TextDisplayForm("\u4E00\u89A7\u8868\u793A", workbookName, content))
                {
                    form.ShowDialog();
                }

                _logger.Info(
                    "Workbook custom document properties displayed. workbook="
                    + _excelInteropService.GetWorkbookFullName(workbook)
                    + ", propertyCount="
                    + properties.Count.ToString());
            }
            catch (Exception ex)
            {
                _logger.Error("ShowCustomDocumentProperties failed.", ex);
                MessageBox.Show(
                    "CustomDocumentProperties \u4E00\u89A7\u3092\u8868\u793A\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002\u30ED\u30B0\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                    ProductTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Method: selects and saves SYSTEM_ROOT for the target workbook.
        /// Args: workbook - target workbook.
        /// Returns: none.
        /// Side effects: shows folder picker, updates docprops, saves workbook, shows user messages.
        /// </summary>
        internal void SelectAndSaveSystemRoot(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                MessageBox.Show("\u5BFE\u8C61\u30D6\u30C3\u30AF\u3092\u53D6\u5F97\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string currentSystemRoot = _pathCompatibilityService.NormalizePath(
                    _excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName));
                string selectedPath = SelectFolderPath(
                    "SYSTEM_ROOT \u306B\u8A2D\u5B9A\u3059\u308B\u30D5\u30A9\u30EB\u30C0\u3092\u9078\u629E\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                    currentSystemRoot);
                if (string.IsNullOrWhiteSpace(selectedPath))
                {
                    return;
                }

                // Block: keep SYSTEM_ROOT and template directory aligned to the same root.
                _excelInteropService.SetDocumentProperty(workbook, SystemRootPropertyName, selectedPath);
                _excelInteropService.SetDocumentProperty(
                    workbook,
                    WordTemplateDirectoryPropertyName,
                    _pathCompatibilityService.CombinePath(selectedPath, TemplateFolderName));
                workbook.Save();

                _logger.Info(
                    "Workbook SYSTEM_ROOT updated manually. workbook="
                    + _excelInteropService.GetWorkbookFullName(workbook)
                    + ", systemRoot="
                    + selectedPath);

                MessageBox.Show(
                    "SYSTEM_ROOT \u3092\u66F4\u65B0\u3057\u307E\u3057\u305F\u3002\r\n" + selectedPath,
                    ProductTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger.Error("SelectAndSaveSystemRoot failed.", ex);
                MessageBox.Show(
                    "SYSTEM_ROOT \u3092\u66F4\u65B0\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002\u30ED\u30B0\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                    ProductTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Method: copies shSample column B values into shHOME column B for the target workbook.
        /// Args: workbook - target Base/Case workbook.
        /// Returns: none.
        /// Side effects: updates worksheet values and shows a completion or error message.
        /// </summary>
        internal void CopySampleColumnBToHome(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                MessageBox.Show("\u5BFE\u8C61\u30D6\u30C3\u30AF\u3092\u53D6\u5F97\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Excel.Worksheet sourceWorksheet = null;
            Excel.Worksheet targetWorksheet = null;
            Excel.Range sourceRange = null;
            Excel.Range targetRange = null;
            Excel.Range clearRange = null;

            try
            {
                sourceWorksheet = _excelInteropService.FindWorksheetByCodeName(workbook, SourceSheetCodeName);
                targetWorksheet = _excelInteropService.FindWorksheetByCodeName(workbook, TargetSheetCodeName);
                if (sourceWorksheet == null || targetWorksheet == null)
                {
                    MessageBox.Show(
                        "shSample \u307E\u305F\u306F shHOME \u30B7\u30FC\u30C8\u3092\u53D6\u5F97\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F\u3002BASE / CASE \u30D6\u30C3\u30AF\u3067\u5B9F\u884C\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                        ProductTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                int sourceLastRow = FindLastUsedRowInColumn(sourceWorksheet, TargetColumnIndex);
                int targetLastRow = FindLastUsedRowInColumn(targetWorksheet, TargetColumnIndex);
                if (sourceLastRow == 0)
                {
                    if (targetLastRow > 0)
                    {
                        clearRange = targetWorksheet.Range[targetWorksheet.Cells[1, TargetColumnIndex], targetWorksheet.Cells[targetLastRow, TargetColumnIndex]];
                        clearRange.ClearContents();
                    }

                    MessageBox.Show(
                        "shSample \u306E B \u5217\u306B\u8EE2\u8A18\u5BFE\u8C61\u306E\u6587\u5B57\u304C\u3042\u308A\u307E\u305B\u3093\u3002shHOME \u306E B \u5217\u306F\u30AF\u30EA\u30A2\u3057\u307E\u3057\u305F\u3002",
                        ProductTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                sourceRange = sourceWorksheet.Range[sourceWorksheet.Cells[1, TargetColumnIndex], sourceWorksheet.Cells[sourceLastRow, TargetColumnIndex]];
                targetRange = targetWorksheet.Range[targetWorksheet.Cells[1, TargetColumnIndex], targetWorksheet.Cells[sourceLastRow, TargetColumnIndex]];
                targetRange.Value2 = sourceRange.Value2;

                if (targetLastRow > sourceLastRow)
                {
                    clearRange = targetWorksheet.Range[targetWorksheet.Cells[sourceLastRow + 1, TargetColumnIndex], targetWorksheet.Cells[targetLastRow, TargetColumnIndex]];
                    clearRange.ClearContents();
                }

                _logger.Info(
                    "Sample column B copied to shHOME. workbook="
                    + _excelInteropService.GetWorkbookFullName(workbook)
                    + ", rowCount="
                    + sourceLastRow.ToString());

                MessageBox.Show(
                    "shSample \u306E B1:B" + sourceLastRow.ToString() + " \u3092 shHOME \u306E B \u5217\u3078\u8EE2\u8A18\u3057\u307E\u3057\u305F\u3002",
                    ProductTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger.Error("CopySampleColumnBToHome failed.", ex);
                MessageBox.Show(
                    "shSample \u304B\u3089 shHOME \u3078\u306E\u8EE2\u8A18\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002\u30ED\u30B0\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                    ProductTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(clearRange);
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(targetRange);
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(sourceRange);
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(targetWorksheet);
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(sourceWorksheet);
            }
        }

        /// <summary>
        /// Method: builds display text for CustomDocumentProperties.
        /// Args: properties - properties to display.
        /// Returns: formatted text for the dialog.
        /// Side effects: none.
        /// </summary>
        private static string BuildCustomDocumentPropertyText(IReadOnlyList<KeyValuePair<string, string>> properties)
        {
            if (properties == null || properties.Count == 0)
            {
                return "CustomDocumentProperties \u306F\u767B\u9332\u3055\u308C\u3066\u3044\u307E\u305B\u3093\u3002";
            }

            var builder = new StringBuilder();
            for (int index = 0; index < properties.Count; index++)
            {
                KeyValuePair<string, string> property = properties[index];
                builder.Append(index + 1);
                builder.Append(". ");
                builder.Append(property.Key ?? string.Empty);
                builder.Append(" = ");
                builder.AppendLine(property.Value ?? string.Empty);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Method: opens a folder picker for SYSTEM_ROOT.
        /// Args: dialogTitle - title, initialDirectory - initial folder.
        /// Returns: selected folder or empty string when cancelled.
        /// Side effects: shows the folder picker dialog.
        /// </summary>
        private string SelectFolderPath(string dialogTitle, string initialDirectory)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Multiselect = false;
                dialog.Title = dialogTitle ?? string.Empty;
                dialog.EnsurePathExists = true;
                dialog.AllowNonFileSystemItems = false;

                string normalizedDirectory = _pathCompatibilityService.NormalizePath(initialDirectory);
                if (!string.IsNullOrWhiteSpace(normalizedDirectory) && System.IO.Directory.Exists(normalizedDirectory))
                {
                    dialog.InitialDirectory = normalizedDirectory;
                    dialog.DefaultDirectory = normalizedDirectory;
                }

                return dialog.ShowDialog() == CommonFileDialogResult.Ok
                    ? _pathCompatibilityService.NormalizePath(dialog.FileName)
                    : string.Empty;
            }
        }

        private static int FindLastUsedRowInColumn(Excel.Worksheet worksheet, int columnIndex)
        {
            if (worksheet == null || columnIndex <= 0)
            {
                return 0;
            }

            Excel.Range lastCell = null;
            Excel.Range firstCell = null;
            try
            {
                lastCell = ((dynamic)worksheet.Cells[worksheet.Rows.Count, columnIndex]).End[Excel.XlDirection.xlUp];
                int lastRow = lastCell == null ? 0 : Convert.ToInt32(lastCell.Row);
                if (lastRow <= 0)
                {
                    return 0;
                }

                firstCell = worksheet.Cells[1, columnIndex] as Excel.Range;
                string firstValue = Convert.ToString(firstCell?.Value2) ?? string.Empty;
                return lastRow == 1 && string.IsNullOrWhiteSpace(firstValue) ? 0 : lastRow;
            }
            finally
            {
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(firstCell);
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(lastCell);
            }
        }

    }
}
