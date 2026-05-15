using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Domain;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class AccountingWorkbookService
	{
		private readonly Application _application;

		private readonly ExcelValidationService _excelValidationService;

		private readonly Logger _logger;

		internal AccountingWorkbookService (Application application, ExcelValidationService excelValidationService, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_excelValidationService = excelValidationService ?? throw new ArgumentNullException ("excelValidationService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal Workbook OpenInCurrentApplication (string workbookPath)
		{
			if (string.IsNullOrWhiteSpace (workbookPath)) {
				throw new ArgumentException ("Workbook path is required.", "workbookPath");
			}
			_logger.Info ("Accounting workbook open in current application. path=" + workbookPath);
			Workbook workbook = _application.Workbooks.Open (workbookPath, 0, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
			Workbook activeWorkbook = _application.ActiveWorkbook;
			_logger.Info ("Accounting workbook open completed. workbook=" + (workbook == null ? string.Empty : (workbook.FullName ?? workbook.Name ?? string.Empty)) + ", activeWorkbook=" + (activeWorkbook == null ? string.Empty : (activeWorkbook.FullName ?? activeWorkbook.Name ?? string.Empty)));
			return workbook;
		}

		internal Workbook OpenReadOnlyInCurrentApplication (string workbookPath)
		{
			if (string.IsNullOrWhiteSpace (workbookPath)) {
				throw new ArgumentException ("Workbook path is required.", "workbookPath");
			}
			_logger.Info ("Accounting workbook open read-only in current application. path=" + workbookPath);
			return _application.Workbooks.Open (workbookPath, 0, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
		}

		internal Workbook OpenReadOnlyHiddenInCurrentApplication (string workbookPath)
		{
			Workbook workbook = OpenReadOnlyInCurrentApplication (workbookPath);
			SetWorkbookWindowsVisible (workbook, visible: false);
			return workbook;
		}

		internal void CloseWithoutSaving (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave (workbook, _logger, "AccountingWorkbookService.CloseWithoutSaving");
		}

		internal void SetWorkbookWindowsVisible (Workbook workbook, bool visible)
		{
			if (workbook == null) {
				return;
			}
			Windows windows = null;
			try {
				windows = workbook.Windows;
				int windowCount = (windows == null) ? 0 : windows.Count;
				for (int index = 1; index <= windowCount; index++) {
					Window window = null;
					try {
						window = windows [index];
						if (window != null) {
							window.Visible = visible;
						}
					} finally {
						ComObjectReleaseService.Release (window);
					}
				}
			} finally {
				ComObjectReleaseService.Release (windows);
			}
			Workbook activeWorkbook = _application.ActiveWorkbook;
			_logger.Info ("Accounting workbook windows visibility updated. workbook=" + (workbook.FullName ?? workbook.Name ?? string.Empty) + ", visible=" + visible + ", activeWorkbook=" + (activeWorkbook == null ? string.Empty : (activeWorkbook.FullName ?? activeWorkbook.Name ?? string.Empty)));
		}

		internal void SaveAsMacroEnabled (Workbook workbook, string savePath)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (string.IsNullOrWhiteSpace (savePath)) {
				throw new ArgumentException ("Save path is required.", "savePath");
			}
			using (var excelApplicationStateScope = new ExcelApplicationStateScope (_application)) {
				excelApplicationStateScope.SetDisplayAlerts (false);
				excelApplicationStateScope.SetEnableEvents (false);
				excelApplicationStateScope.SetScreenUpdating (false);
				XlFileFormat saveFormat = ResolveSaveAsFileFormat (savePath);
				workbook.SaveAs (savePath, saveFormat, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				_logger.Info ("Accounting workbook SaveAs completed. path=" + savePath + ", format=" + saveFormat);
			}
		}

		internal void WriteCell (Workbook workbook, string sheetName, string address, string valueText)
		{
			WriteCellValue (workbook, sheetName, address, valueText ?? string.Empty);
		}

		internal void WriteCellValue (Workbook workbook, string sheetName, string address, object value)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.Value2 = value ?? string.Empty;
				_logger.Debug ("AccountingWorkbookService", "WriteCell sheet=" + sheetName + ", address=" + address + ", value=" + (Convert.ToString (value) ?? string.Empty));
			} catch (Exception innerException) {
				throw new InvalidOperationException ("シート「" + sheetName + "」のセル「" + address + "」への書き込みに失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void WriteCellFormula (Workbook workbook, string sheetName, string address, string formula)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.Formula = formula ?? string.Empty;
				_logger.Debug ("AccountingWorkbookService", "WriteCellFormula sheet=" + sheetName + ", address=" + address);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("セルへの数式書き込みに失敗しました。sheet=" + sheetName + ", address=" + address, innerException);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void WriteSameValueToSheets (Workbook workbook, IEnumerable<string> sheetNames, string address, string valueText)
		{
			if (sheetNames == null) {
				return;
			}
			foreach (string sheetName in sheetNames) {
				WriteCell (workbook, sheetName, address, valueText);
			}
		}

		internal void CopyFormulaRange (Workbook workbook, string sourceSheetName, string sourceAddress, string targetSheetName, string targetAddress)
		{
			Worksheet worksheet = null;
			Worksheet worksheet2 = null;
			Range range = null;
			Range range2 = null;
			try {
				worksheet = GetWorksheet (workbook, sourceSheetName);
				worksheet2 = GetWorksheet (workbook, targetSheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)sourceAddress, Type.Missing);
				range2 = ((_Worksheet)worksheet2).get_Range ((object)targetAddress, Type.Missing);
				range2.Formula = range.Formula;
				_logger.Debug ("AccountingWorkbookService", "CopyFormulaRange source=" + sourceSheetName + "!" + sourceAddress + ", target=" + targetSheetName + "!" + targetAddress);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("数式範囲コピーに失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet2);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void CopyValueRange (Workbook workbook, string sourceSheetName, string sourceAddress, string targetSheetName, string targetAddress)
		{
			Worksheet worksheet = null;
			Worksheet worksheet2 = null;
			Range range = null;
			Range range2 = null;
			try {
				worksheet = GetWorksheet (workbook, sourceSheetName);
				worksheet2 = GetWorksheet (workbook, targetSheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)sourceAddress, Type.Missing);
				range2 = ((_Worksheet)worksheet2).get_Range ((object)targetAddress, Type.Missing);
				range2.Value2 = range.Value2;
				_logger.Debug ("AccountingWorkbookService", "CopyValueRange source=" + sourceSheetName + "!" + sourceAddress + ", target=" + targetSheetName + "!" + targetAddress);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("値範囲コピーに失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet2);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ClearMergeAreaContents (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			Range range2 = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range2 = range.MergeArea;
				range2.ClearContents ();
				_logger.Debug ("AccountingWorkbookService", "ClearMergeAreaContents sheet=" + sheetName + ", address=" + address);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("マージ領域クリアに失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void SetInteriorColorIndex (Workbook workbook, string sheetName, string address, int colorIndex)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.Interior.ColorIndex = colorIndex;
				_logger.Debug ("AccountingWorkbookService", "SetInteriorColorIndex sheet=" + sheetName + ", address=" + address + ", colorIndex=" + colorIndex);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("セル色設定に失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void UnprotectSheet (Workbook workbook, string sheetName)
		{
			Worksheet worksheet = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				worksheet.Unprotect (Type.Missing);
				_logger.Debug ("AccountingWorkbookService", "UnprotectSheet sheet=" + sheetName);
			} finally {
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ProtectSheetUiOnly (Workbook workbook, string sheetName)
		{
			Worksheet worksheet = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				worksheet.Protect (Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				_logger.Debug ("AccountingWorkbookService", "ProtectSheetUiOnly sheet=" + sheetName);
			} finally {
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal IDisposable BeginInitializationScope ()
		{
			var excelApplicationStateScope = new ExcelApplicationStateScope (_application);
			excelApplicationStateScope.SetScreenUpdating (false);
			excelApplicationStateScope.SetEnableEvents (false);
			return excelApplicationStateScope;
		}

		internal bool EnsureNumberFormatLocal (Workbook workbook, string sheetName, string address, string numberFormatLocal)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				string a = Convert.ToString ((dynamic)range.NumberFormatLocal) ?? string.Empty;
				string text = numberFormatLocal ?? string.Empty;
				if (string.Equals (a, text, StringComparison.Ordinal)) {
					_logger.Debug ("AccountingWorkbookService", "EnsureNumberFormatLocal skipped. sheet=" + sheetName + ", address=" + address);
					return false;
				}
				range.NumberFormatLocal = text;
				_logger.Debug ("AccountingWorkbookService", "EnsureNumberFormatLocal updated. sheet=" + sheetName + ", address=" + address);
				return true;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal object ReadCellValue (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				return range.Value2;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal object ReadRangeFormulaSnapshot (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				return range.Formula;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void WriteRangeFormulaSnapshot (Workbook workbook, string sheetName, string address, object formulaSnapshot)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.Formula = formulaSnapshot ?? string.Empty;
				_logger.Debug ("AccountingWorkbookService", "WriteRangeFormulaSnapshot sheet=" + sheetName + ", address=" + address);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void WriteCellValue (Worksheet worksheet, string address, object value)
		{
			Range range = null;
			try {
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.Value2 = value;
			} finally {
				ComObjectReleaseService.Release (range);
			}
		}

		internal void WriteRangeValues (Worksheet worksheet, string address, object[,] values)
		{
			Range range = null;
			try {
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.Value2 = values;
			} finally {
				ComObjectReleaseService.Release (range);
			}
		}

		internal object ReadCellValue (Worksheet worksheet, string address)
		{
			Range range = null;
			try {
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				return range.Value2;
			} finally {
				ComObjectReleaseService.Release (range);
			}
		}

		internal void ClearRangeContents (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.ClearContents ();
				_logger.Debug ("AccountingWorkbookService", "ClearRangeContents sheet=" + sheetName + ", address=" + address);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ClearValidation (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			Range targetRange = null;
			Validation validation = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				targetRange = range.MergeArea;
				validation = targetRange.Validation;
				validation.Delete ();
				_logger.Debug ("AccountingWorkbookService", "ClearValidation sheet=" + sheetName + ", address=" + address);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("シート「" + sheetName + "」のセル「" + address + "」の入力規則解除に失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (validation);
				ComObjectReleaseService.Release (targetRange);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void WriteNamedRangeValue (Workbook workbook, string sheetName, string rangeName, object value)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				range.Value2 = value ?? string.Empty;
				_logger.Debug ("AccountingWorkbookService", "WriteNamedRangeValue sheet=" + sheetName + ", rangeName=" + rangeName);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal string ReadDisplayTextByNamedRange (Workbook workbook, string sheetName, string rangeName)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				return (Convert.ToString ((dynamic)range.Text) ?? string.Empty).Trim ();
			} catch {
				return string.Empty;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal object ReadNamedRangeValue (Workbook workbook, string sheetName, string rangeName)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				return range.Value2;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal string GetNamedRangeAddress (Workbook workbook, string sheetName, string rangeName)
		{
			Worksheet worksheet = null;
			Range range = null;
			Worksheet ownerWorksheet = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				ownerWorksheet = range.Worksheet as Worksheet;
				string ownerSheetName = ownerWorksheet == null ? string.Empty : (ownerWorksheet.Name ?? string.Empty);
				if (!string.Equals (ownerSheetName, sheetName, StringComparison.Ordinal)) {
					throw new InvalidOperationException ("名前定義の参照先シートが想定と異なります。sheet=" + sheetName + ", name=" + rangeName + ", actualSheet=" + ownerSheetName);
				}
				return range.get_Address ((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("名前定義が見つからないか参照できません。sheet=" + sheetName + ", name=" + rangeName, innerException);
			} finally {
				ComObjectReleaseService.Release (ownerWorksheet);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ClearNamedRangeContents (Workbook workbook, string sheetName, string rangeName)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				range.ClearContents ();
				_logger.Debug ("AccountingWorkbookService", "ClearNamedRangeContents sheet=" + sheetName + ", rangeName=" + rangeName);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ClearNamedRangeMergeAreaContents (Workbook workbook, string sheetName, string rangeName)
		{
			Worksheet worksheet = null;
			Range range = null;
			Range range2 = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				range2 = range.MergeArea;
				range2.ClearContents ();
				_logger.Debug ("AccountingWorkbookService", "ClearNamedRangeMergeAreaContents sheet=" + sheetName + ", rangeName=" + rangeName);
			} finally {
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void SetPrintAreaByBounds (Workbook workbook, string sheetName, int lastRow, int lastColumn)
		{
			Worksheet worksheet = null;
			Range range = null;
			Range range2 = null;
			Range range3 = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = worksheet.Cells [1, 1] as Range;
				range2 = worksheet.Cells [lastRow, lastColumn] as Range;
				range3 = ((_Worksheet)worksheet).get_Range ((object)range, (object)range2);
				worksheet.PageSetup.PrintArea = range3.get_Address ((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
				_logger.Debug ("AccountingWorkbookService", "SetPrintAreaByBounds sheet=" + sheetName + ", lastRow=" + lastRow + ", lastColumn=" + lastColumn);
			} finally {
				ComObjectReleaseService.Release (range3);
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal bool HasListObjectRange (Workbook workbook, string sheetName, string expectedAddress)
		{
			Worksheet worksheet = null;
			ListObjects listObjects = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				listObjects = worksheet.ListObjects;
				int count = listObjects == null ? 0 : listObjects.Count;
				for (int index = 1; index <= count; index++) {
					ListObject listObject = null;
					Range range = null;
					try {
						listObject = listObjects [index];
						range = listObject.Range;
						string address = range.get_Address ((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
						if (string.Equals (address, expectedAddress, StringComparison.OrdinalIgnoreCase)) {
							return true;
						}
					} finally {
						ComObjectReleaseService.Release (range);
						ComObjectReleaseService.Release (listObject);
					}
				}
				return false;
			} finally {
				ComObjectReleaseService.Release (listObjects);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal bool SortListObjectAscendingByColumn (Workbook workbook, string sheetName, string tableRangeAddress, int keyColumnIndex)
		{
			if (keyColumnIndex <= 0) {
				throw new ArgumentOutOfRangeException ("keyColumnIndex");
			}
			Worksheet worksheet = null;
			ListObjects listObjects = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				listObjects = worksheet.ListObjects;
				int count = listObjects == null ? 0 : listObjects.Count;
				for (int index = 1; index <= count; index++) {
					ListObject listObject = null;
					Range tableRange = null;
					try {
						listObject = listObjects [index];
						tableRange = listObject.Range;
						string address = tableRange.get_Address ((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
						if (!string.Equals (address, tableRangeAddress, StringComparison.OrdinalIgnoreCase)) {
							continue;
						}
						SortListObjectAscending (listObject, keyColumnIndex);
						_logger.Debug ("AccountingWorkbookService", "SortListObjectAscendingByColumn sheet=" + sheetName + ", table=" + (listObject.Name ?? string.Empty) + ", range=" + tableRangeAddress + ", keyColumnIndex=" + keyColumnIndex.ToString (System.Globalization.CultureInfo.InvariantCulture));
						return true;
					} finally {
						ComObjectReleaseService.Release (tableRange);
						ComObjectReleaseService.Release (listObject);
					}
				}
				return false;
			} finally {
				ComObjectReleaseService.Release (listObjects);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		private static void SortListObjectAscending (ListObject listObject, int keyColumnIndex)
		{
			ListColumn keyColumn = null;
			Range keyRange = null;
			Sort sort = null;
			SortFields sortFields = null;
			SortField sortField = null;
			try {
				keyColumn = listObject.ListColumns [keyColumnIndex];
				keyRange = keyColumn.Range;
				sort = listObject.Sort;
				sortFields = sort.SortFields;
				sortFields.Clear ();
				sortField = sortFields.Add (keyRange, XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, Type.Missing, XlSortDataOption.xlSortNormal);
				sort.Header = XlYesNoGuess.xlYes;
				sort.MatchCase = false;
				sort.Apply ();
			} finally {
				ComObjectReleaseService.Release (sortField);
				ComObjectReleaseService.Release (sortFields);
				ComObjectReleaseService.Release (sort);
				ComObjectReleaseService.Release (keyRange);
				ComObjectReleaseService.Release (keyColumn);
			}
		}

		internal void SetRowHidden (Workbook workbook, string sheetName, int rowNumber, bool hidden)
		{
			Worksheet worksheet = null;
			Range row = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				row = worksheet.Rows [rowNumber, Type.Missing] as Range;
				if (row != null) {
					row.Hidden = hidden;
				}
			} finally {
				ComObjectReleaseService.Release (row);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void SetColumnHidden (Workbook workbook, string sheetName, string columnAddress, bool hidden)
		{
			Worksheet worksheet = null;
			Range column = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				column = worksheet.Columns [columnAddress, Type.Missing] as Range;
				if (column != null) {
					column.Hidden = hidden;
				}
			} finally {
				ComObjectReleaseService.Release (column);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void SetHorizontalAlignmentCenter (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
				_logger.Debug ("AccountingWorkbookService", "SetHorizontalAlignmentCenter sheet=" + sheetName + ", address=" + address);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void SortRangeAscending (Workbook workbook, string sheetName, string sortRangeAddress, string keyAddress)
		{
			Worksheet worksheet = null;
			Range range = null;
			Range range2 = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)sortRangeAddress, Type.Missing);
				range2 = ((_Worksheet)worksheet).get_Range ((object)keyAddress, Type.Missing);
				range.Sort (range2, XlSortOrder.xlAscending, Type.Missing, Type.Missing, XlSortOrder.xlAscending, Type.Missing, XlSortOrder.xlAscending, XlYesNoGuess.xlNo, Type.Missing, Type.Missing);
				_logger.Debug ("AccountingWorkbookService", "SortRangeAscending sheet=" + sheetName + ", range=" + sortRangeAddress + ", key=" + keyAddress);
			} finally {
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal int GetLastUsedRowInColumn (Workbook workbook, string sheetName, string columnAddress)
		{
			Worksheet worksheet = null;
			Range comObject = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				comObject = (dynamic)worksheet.Columns [columnAddress, Type.Missing];
				range = worksheet.Cells [worksheet.Rows.Count, columnAddress] as Range;
				return range?.get_End (XlDirection.xlUp).Row ?? 1;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (comObject);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal string ReadNamedRangeText (Workbook workbook, string sheetName, string rangeName)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ResolveNamedRange (workbook, worksheet, rangeName);
				return Convert.ToString ((dynamic)range.Value2) ?? string.Empty;
			} catch {
				return string.Empty;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal Range TryGetActiveCell (Workbook workbook)
		{
			try {
				Range range = (workbook?.Application)?.ActiveCell;
				return (range != null && range.Cells.Count == 1) ? range : null;
			} catch {
				return null;
			}
		}

		internal bool IsWithinRange (Workbook workbook, string sheetName, Range cell, string allowedAddress)
		{
			Worksheet worksheet = null;
			Range range = null;
			Range range2 = null;
			try {
				if (cell == null) {
					return false;
				}
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)allowedAddress, Type.Missing);
				range2 = worksheet.Application.Intersect (cell, range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				return range2 != null;
			} finally {
				ComObjectReleaseService.Release (range2);
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ExecuteGoalSeek (Workbook workbook, string sheetName, string formulaCellAddress, Range changingCell, double goalValue)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)formulaCellAddress, Type.Missing);
				range.GoalSeek (goalValue, changingCell);
				_logger.Info ("Accounting goal seek executed. sheet=" + sheetName + ", formulaCell=" + formulaCellAddress + ", target=" + goalValue);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("逆算に失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ExecuteGoalSeek (Workbook workbook, string sheetName, string formulaCellAddress, string changingCellAddress, double goalValue)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)changingCellAddress, Type.Missing);
				ExecuteGoalSeek (workbook, sheetName, formulaCellAddress, range, goalValue);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ExecuteGoalSeekOrThrow (Workbook workbook, string sheetName, string formulaCellAddress, string changingCellAddress, double goalValue)
		{
			AccountingGoalSeekExecutionResult result = ExecuteGoalSeekAndReadResult (workbook, sheetName, formulaCellAddress, changingCellAddress, goalValue);
			if (!result.Succeeded) {
				throw new InvalidOperationException (
					"GoalSeek に失敗しました。sheet=" + sheetName +
					", formulaCell=" + formulaCellAddress +
					", changingCell=" + changingCellAddress +
					", target=" + Convert.ToString (goalValue, CultureInfo.InvariantCulture) +
					", current=" + ConvertGoalSeekValueToString (result.CurrentValue) +
					", goalSeekResult=False");
			}
		}

		internal AccountingGoalSeekExecutionResult ExecuteGoalSeekAndReadResult (Workbook workbook, string sheetName, string formulaCellAddress, string changingCellAddress, double goalValue)
		{
			Worksheet worksheet = null;
			Range formulaCell = null;
			Range changingCell = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				formulaCell = ((_Worksheet)worksheet).get_Range ((object)formulaCellAddress, Type.Missing);
				changingCell = ((_Worksheet)worksheet).get_Range ((object)changingCellAddress, Type.Missing);
				bool succeeded = formulaCell.GoalSeek (goalValue, changingCell);
				object currentValue = formulaCell.Value2;
				_logger.Info ("Accounting goal seek executed. sheet=" + sheetName + ", formulaCell=" + formulaCellAddress + ", changingCell=" + changingCellAddress + ", target=" + Convert.ToString (goalValue, CultureInfo.InvariantCulture) + ", current=" + ConvertGoalSeekValueToString (currentValue) + ", goalSeekResult=" + succeeded.ToString ());
				return new AccountingGoalSeekExecutionResult (succeeded, currentValue);
			} catch (Exception innerException) {
				throw new InvalidOperationException (
					"GoalSeek 実行中に例外が発生しました。sheet=" + sheetName +
					", formulaCell=" + formulaCellAddress +
					", changingCell=" + changingCellAddress +
					", target=" + Convert.ToString (goalValue, CultureInfo.InvariantCulture) +
					", exceptionType=" + innerException.GetType ().FullName +
					", exceptionMessage=" + innerException.Message,
					innerException);
			} finally {
				ComObjectReleaseService.Release (changingCell);
				ComObjectReleaseService.Release (formulaCell);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		private static string ConvertGoalSeekValueToString (object value)
		{
			return Convert.ToString (value, CultureInfo.InvariantCulture) ?? string.Empty;
		}

		internal void RoundDownCell (Range cell, int digits)
		{
			if (cell == null) {
				return;
			}
			try {
				double num = Convert.ToDouble ((dynamic)cell.Value2);
				double num3 = RoundDownValue (num, digits);
				cell.Value2 = num3;
			} catch (Exception innerException) {
				throw new InvalidOperationException ("逆算結果の丸めに失敗しました。", innerException);
			}
		}

		internal void RoundDownCell (Workbook workbook, string sheetName, string address, int digits)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				RoundDownCell (range, digits);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		private static double RoundDownValue (double value, int digits)
		{
			double scale = Math.Pow (10.0, digits);
			return Math.Floor (value * scale) / scale;
		}

		internal double ReadDouble (Range cell)
		{
			try {
				return (cell == null) ? ((object)0.0) : Convert.ToDouble ((dynamic)cell.Value2);
			} catch {
				return 0.0;
			}
		}

		internal DateTime ReadDateCell (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				object value = range.Value2;
				if (value is double) {
					return DateTime.FromOADate ((double)value);
				}
				return Convert.ToDateTime (value);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("日付セルの読み取りに失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal string ReadText (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				return (Convert.ToString ((dynamic)range.Value2) ?? string.Empty).Trim ();
			} catch {
				return string.Empty;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal string ReadDisplayText (Workbook workbook, string sheetName, string address)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				return (Convert.ToString ((dynamic)range.Text) ?? string.Empty).Trim ();
			} catch {
				return string.Empty;
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal AccountingLawyerMappingResult ReflectLawyers (Workbook workbook, string lawyerLinesText)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			List<string> list = ExtractLawyerLines (lawyerLinesText);
			AccountingLawyerMappingResult accountingLawyerMappingResult = new AccountingLawyerMappingResult ();
			int num = Math.Min (list.Count, AccountingSetSpec.MaximumLawyerCount);
			for (int i = 0; i < AccountingSetSpec.LawyerReflectionTargets.Count; i++) {
				Worksheet worksheet = null;
				AccountingLawyerReflectionTarget target = AccountingSetSpec.LawyerReflectionTargets [i];
				try {
					worksheet = GetWorksheet (workbook, target.SheetName);
					for (int j = 0; j < AccountingSetSpec.MaximumLawyerCount; j++) {
						Range range = null;
						try {
							range = ((_Worksheet)worksheet).get_Range ((object)target.StartCellAddress, Type.Missing).get_Offset ((object)j, (object)0);
							range.Value2 = string.Empty;
							if (j < num) {
								string text = FindFirstValidationMatch (range, list [j]);
								if (text.Length > 0) {
									range.Value2 = text;
									accountingLawyerMappingResult.AssignedCount++;
									_logger.Debug ("AccountingWorkbookService", "Lawyer matched. sheet=" + target.SheetName + ", startCell=" + target.StartCellAddress + ", rowOffset=" + j + ", source=" + list [j] + ", matched=" + text);
								} else {
									accountingLawyerMappingResult.MissingMatchCount++;
									_logger.Warn ("Accounting lawyer match was not found. sheet=" + target.SheetName + ", startCell=" + target.StartCellAddress + ", rowOffset=" + j + ", source=" + list [j]);
								}
							}
						} finally {
							ComObjectReleaseService.Release (range);
						}
					}
				} finally {
					ComObjectReleaseService.Release (worksheet);
				}
			}
			if (list.Count > AccountingSetSpec.MaximumLawyerCount) {
				accountingLawyerMappingResult.OverflowCount = list.Count - AccountingSetSpec.MaximumLawyerCount;
			}
			_logger.Info ("Accounting lawyers reflected. assigned=" + accountingLawyerMappingResult.AssignedCount + ", missingMatch=" + accountingLawyerMappingResult.MissingMatchCount + ", overflow=" + accountingLawyerMappingResult.OverflowCount);
			return accountingLawyerMappingResult;
		}

		internal void ActivateInvoiceEntry (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, "請求書");
				range = ((_Worksheet)worksheet).get_Range ((object)"A3", Type.Missing);
				workbook.Activate ();
				Workbook activeWorkbook = _application.ActiveWorkbook;
				_logger.Info ("Accounting invoice entry workbook activated. workbook=" + (workbook.FullName ?? workbook.Name ?? string.Empty) + ", activeWorkbook=" + (activeWorkbook == null ? string.Empty : (activeWorkbook.FullName ?? activeWorkbook.Name ?? string.Empty)));
				worksheet.Activate ();
				_logger.Info ("Accounting invoice entry worksheet activated. workbook=" + (workbook.FullName ?? workbook.Name ?? string.Empty) + ", sheet=" + (worksheet == null ? string.Empty : (worksheet.CodeName ?? worksheet.Name ?? string.Empty)));
				try {
					range.Select ();
					_logger.Info ("Accounting invoice entry activated. sheet=請求書, address=A3, selected=True");
				} catch (COMException ex) {
					_logger.Warn ("Accounting invoice entry select skipped after activation. sheet=請求書, address=A3, message=" + ex.Message);
				}
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void ActivateCell (Workbook workbook, string sheetName, string address)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, sheetName);
				range = ((_Worksheet)worksheet).get_Range ((object)address, Type.Missing);
				workbook.Activate ();
				worksheet.Activate ();
				range.Select ();
				_logger.Debug ("AccountingWorkbookService", "ActivateCell sheet=" + sheetName + ", address=" + address);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal Worksheet GetWorksheet (Workbook workbook, string sheetName)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (string.IsNullOrWhiteSpace (sheetName)) {
				throw new ArgumentException ("Sheet name is required.", "sheetName");
			}
			if (!(workbook.Worksheets [sheetName] is Worksheet result)) {
				throw new InvalidOperationException ("シートが見つかりません: " + sheetName);
			}
			return result;
		}

		internal void HighlightAccountingImportTargets (Workbook workbook)
		{
			SetAccountingImportTargetHighlight (workbook, 36);
		}

		internal void ClearAccountingImportTargetHighlight (Workbook workbook)
		{
			SetAccountingImportTargetHighlight (workbook, 0);
		}

		private void SetAccountingImportTargetHighlight (Workbook workbook, int colorIndex)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = GetWorksheet (workbook, "会計依頼書");
				range = ((_Worksheet)worksheet).get_Range ((object)"F15:F20", Type.Missing);
				range.Interior.ColorIndex = colorIndex;
				_logger.Info ("Accounting import target range highlight changed. sheet=会計依頼書, address=F15:F20, colorIndex=" + colorIndex);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("会計依頼書の黄色エリア表示に失敗しました。", innerException);
			} finally {
				ComObjectReleaseService.Release (range);
				ComObjectReleaseService.Release (worksheet);
			}
		}

		internal void HighlightReverseToolTargets (Workbook workbook, string sheetName)
		{
			SetInteriorColorIndex (workbook, sheetName, "F17", 36);
			SetInteriorColorIndex (workbook, sheetName, "F18", 36);
			SetInteriorColorIndex (workbook, sheetName, "F19", 36);
			SetInteriorColorIndex (workbook, sheetName, "F20", 36);
		}

		internal void ClearReverseToolTargets (Workbook workbook, string sheetName)
		{
			SetInteriorColorIndex (workbook, sheetName, "F17", 0);
			SetInteriorColorIndex (workbook, sheetName, "F18", 0);
			SetInteriorColorIndex (workbook, sheetName, "F19", 0);
			SetInteriorColorIndex (workbook, sheetName, "F20", 0);
		}

		private string FindFirstValidationMatch (Range targetCell, string keyword)
		{
			IReadOnlyList<string> validationCandidates = _excelValidationService.GetValidationCandidates (targetCell);
			string text = NormalizeMatchText (keyword);
			if (text.Length == 0) {
				return string.Empty;
			}
			for (int i = 0; i < validationCandidates.Count; i++) {
				string text2 = validationCandidates [i] ?? string.Empty;
				if (text2.Length != 0 && NormalizeMatchText (text2).IndexOf (text, StringComparison.OrdinalIgnoreCase) >= 0) {
					return text2;
				}
			}
			return string.Empty;
		}

		private static List<string> ExtractLawyerLines (string sourceText)
		{
			List<string> list = new List<string> ();
			string text = (sourceText ?? string.Empty).Replace ("\r\n", "\n").Replace ('\r', '\n');
			string[] array = text.Split ('\n');
			for (int i = 0; i < array.Length; i++) {
				string text2 = (array [i] ?? string.Empty).Trim ();
				if (text2.Length > 0) {
					list.Add (text2);
				}
			}
			return list;
		}

		private static string NormalizeMatchText (string text)
		{
			return (text ?? string.Empty).Trim ().Replace (" ", string.Empty).Replace ("\u3000", string.Empty);
		}

		private Range ResolveNamedRange (Workbook workbook, Worksheet worksheet, string rangeName)
		{
			if (worksheet == null) {
				throw new ArgumentNullException ("worksheet");
			}
			if (string.IsNullOrWhiteSpace (rangeName)) {
				throw new ArgumentException ("Range name is required.", "rangeName");
			}
			Name name = null;
			Name name2 = null;
			try {
				try {
					name = worksheet.Names.Item (rangeName, Type.Missing, Type.Missing);
					if (name != null) {
						return name.RefersToRange;
					}
				} catch {
				}
				try {
					name = worksheet.Names.Item ((worksheet.Name ?? string.Empty) + "!" + rangeName, Type.Missing, Type.Missing);
					if (name != null) {
						return name.RefersToRange;
					}
				} catch {
				}
				try {
					name2 = workbook?.Names.Item (rangeName, Type.Missing, Type.Missing);
					if (name2 != null) {
						return name2.RefersToRange;
					}
				} catch {
				}
				return ((_Worksheet)worksheet).get_Range ((object)rangeName, Type.Missing);
			} finally {
				ComObjectReleaseService.Release (name2);
				ComObjectReleaseService.Release (name);
			}
		}

		private static XlFileFormat ResolveSaveAsFileFormat (string savePath)
		{
			string extension = WorkbookFileNameResolver.GetWorkbookExtensionOrDefault (savePath);
			if (string.Equals (extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
				return XlFileFormat.xlOpenXMLWorkbook;
			}
			return XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
		}
	}
}
