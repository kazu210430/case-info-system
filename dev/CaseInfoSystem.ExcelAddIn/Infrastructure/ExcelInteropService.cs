using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Domain;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class ExcelInteropService : IExcelInteropService
	{
		private const string CaseListSheetName = "案件一覧";

		private readonly Application _application;

		private readonly Logger _logger;

		private readonly PathCompatibilityService _pathCompatibilityService;

		internal ExcelInteropService (Application application, Logger logger, PathCompatibilityService pathCompatibilityService)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
		}

		public Workbook GetActiveWorkbook ()
		{
			try {
				return _application.ActiveWorkbook;
			} catch (Exception exception) {
				_logger.Error ("GetActiveWorkbook failed.", exception);
				return null;
			}
		}

		public Window GetActiveWindow ()
		{
			try {
				return _application.ActiveWindow;
			} catch (Exception exception) {
				_logger.Error ("GetActiveWindow failed.", exception);
				return null;
			}
		}

		public string GetWorkbookFullName (Workbook workbook)
		{
			try {
				return (workbook == null) ? string.Empty : (workbook.FullName ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		internal string GetWorkbookName (Workbook workbook)
		{
			try {
				return (workbook == null) ? string.Empty : (workbook.Name ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		internal string GetWorkbookPath (Workbook workbook)
		{
			try {
				return (workbook == null) ? string.Empty : (workbook.Path ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		internal string TryGetDocumentProperty (Workbook workbook, string propertyName)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (propertyName)) {
				return string.Empty;
			}
			object obj = null;
			object obj2 = null;
			try {
				obj = workbook.CustomDocumentProperties;
				dynamic val = obj;
				obj2 = val [propertyName];
				dynamic val2 = obj2;
				object obj3 = val2.Value;
				return Convert.ToString (obj3) ?? string.Empty;
			} catch {
				return string.Empty;
			} finally {
				ReleaseComObject (obj2);
				ReleaseComObject (obj);
			}
		}

		internal void SetDocumentProperty (Workbook workbook, string propertyName, string value)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (string.IsNullOrWhiteSpace (propertyName)) {
				throw new ArgumentException ("Property name is required.", "propertyName");
			}
			object obj = null;
			try {
				obj = workbook.CustomDocumentProperties;
				dynamic val = obj;
				try {
					val [propertyName].Value = value ?? string.Empty;
				} catch {
					val.Add (propertyName, false, 4, value ?? string.Empty);
				}
			} finally {
				ReleaseComObject (obj);
			}
		}

		internal IReadOnlyList<KeyValuePair<string, string>> GetCustomDocumentProperties (Workbook workbook)
		{
			List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>> ();
			if (workbook == null) {
				return list;
			}
			object obj = null;
			try {
				obj = workbook.CustomDocumentProperties;
				if (!(obj is DocumentProperties documentProperties)) {
					return list;
				}
				foreach (DocumentProperty item in documentProperties) {
					try {
						string key = Convert.ToString (item.Name) ?? string.Empty;
						string value = Convert.ToString ((dynamic)item.Value) ?? string.Empty;
						list.Add (new KeyValuePair<string, string> (key, value));
					} finally {
						ReleaseComObject (item);
					}
				}
			} catch (Exception exception) {
				_logger.Error ("GetCustomDocumentProperties failed.", exception);
			} finally {
				ReleaseComObject (obj);
			}
			return list;
		}

		public Window GetFirstVisibleWindow (Workbook workbook)
		{
			if (workbook == null) {
				return null;
			}
			try {
				foreach (Window window in workbook.Windows) {
					if (window != null && window.Visible) {
						return window;
					}
				}
			} catch (Exception exception) {
				_logger.Error ("GetFirstVisibleWindow failed.", exception);
			}
			return null;
		}

		public string GetActiveSheetCodeName (Workbook workbook)
		{
			try {
				Worksheet worksheet = ((workbook == null) ? null : (workbook.ActiveSheet as Worksheet));
				return (worksheet == null) ? string.Empty : (worksheet.CodeName ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		internal Worksheet FindWorksheetByCodeName (Workbook workbook, string sheetCodeName)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (sheetCodeName)) {
				return null;
			}
			try {
				foreach (Worksheet worksheet in workbook.Worksheets) {
					if (string.Equals (worksheet.CodeName, sheetCodeName, StringComparison.OrdinalIgnoreCase)) {
						return worksheet;
					}
				}
			} catch (Exception exception) {
				_logger.Error ("FindWorksheetByCodeName failed.", exception);
			}
			return null;
		}

		internal Worksheet FindWorksheetByName (Workbook workbook, string sheetName)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (sheetName)) {
				return null;
			}
			try {
				return workbook.Worksheets [sheetName] as Worksheet;
			} catch (Exception exception) {
				_logger.Error ("FindWorksheetByName failed. sheetName=" + sheetName, exception);
				return null;
			}
		}

		internal IReadOnlyDictionary<string, string> ReadKeyValueMapFromColumnsAandB (Worksheet worksheet)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			if (worksheet == null) {
				return dictionary;
			}
			Range range = null;
			try {
				int num = ((dynamic)worksheet.Cells [worksheet.Rows.Count, "A"]).End [XlDirection.xlUp].Row;
				if (num < 1) {
					return dictionary;
				}
				range = ((_Worksheet)worksheet).get_Range ((object)"A1", (object)("B" + num));
				if (!(range.Value2 is object[,] array)) {
					return dictionary;
				}
				int upperBound = array.GetUpperBound (0);
				for (int i = 1; i <= upperBound; i++) {
					string text = (Convert.ToString (array [i, 1]) ?? string.Empty).Trim ();
					if (text.Length != 0) {
						string value = Convert.ToString (array [i, 2]) ?? string.Empty;
						dictionary [text] = value;
					}
				}
				return dictionary;
			} catch (Exception exception) {
				_logger.Error ("ReadKeyValueMapFromColumnsAandB failed.", exception);
				return dictionary;
			} finally {
				ReleaseComObject (range);
			}
		}

		internal IReadOnlyList<IReadOnlyDictionary<string, string>> ReadRecordsFromHeaderRow (Worksheet worksheet)
		{
			List<IReadOnlyDictionary<string, string>> list = new List<IReadOnlyDictionary<string, string>> ();
			if (worksheet == null) {
				return list;
			}
			Range range = null;
			try {
				int num = ((dynamic)worksheet.Cells [worksheet.Rows.Count, 1]).End [XlDirection.xlUp].Row;
				int num2 = ((dynamic)worksheet.Cells [1, worksheet.Columns.Count]).End [XlDirection.xlToLeft].Column;
				if (num < 2 || num2 < 1) {
					return list;
				}
				range = worksheet.Range [(dynamic)worksheet.Cells [1, 1], (dynamic)worksheet.Cells [num, num2]];
				if (!(range.Value2 is object[,] array)) {
					return list;
				}
				string[] array2 = new string[num2 + 1];
				for (int i = 1; i <= num2; i++) {
					array2 [i] = (Convert.ToString (array [1, i]) ?? string.Empty).Trim ();
				}
				for (int j = 2; j <= num; j++) {
					Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
					bool flag = false;
					for (int k = 1; k <= num2; k++) {
						string text = array2 [k];
						if (text.Length != 0) {
							string text2 = Convert.ToString (array [j, k]) ?? string.Empty;
							if (text2.Length > 0) {
								flag = true;
							}
							dictionary [text] = text2;
						}
					}
					if (flag) {
						list.Add (dictionary);
					}
				}
				return list;
			} catch (Exception exception) {
				_logger.Error ("ReadRecordsFromHeaderRow failed.", exception);
				return list;
			} finally {
				ReleaseComObject (range);
			}
		}

		internal IReadOnlyDictionary<string, string> ReadFieldValuesFromDefinitions (Worksheet worksheet, IEnumerable<CaseListFieldDefinition> definitions)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			if (worksheet == null || definitions == null) {
				return dictionary;
			}
			Range range = null;
			try {
				int num = 0;
				int num2 = 0;
				List<CaseListFieldDefinition> list = new List<CaseListFieldDefinition> ();
				foreach (CaseListFieldDefinition definition in definitions) {
					if (definition == null || string.IsNullOrWhiteSpace (definition.FieldKey)) {
						continue;
					}
					list.Add (definition);
					if (TryParseCellAddress (definition.SourceCellAddress, out var rowIndex, out var columnIndex)) {
						if (rowIndex > num) {
							num = rowIndex;
						}
						if (columnIndex > num2) {
							num2 = columnIndex;
						}
					}
				}
				object[,] grid = null;
				if (num > 0 && num2 > 0) {
					range = worksheet.Range [(dynamic)worksheet.Cells [1, 1], (dynamic)worksheet.Cells [num, num2]];
					grid = range.Value2 as object[,];
				}
				foreach (CaseListFieldDefinition item in list) {
					string value = ReadValueByDefinition (worksheet, grid, item);
					dictionary [item.FieldKey] = value;
				}
				return dictionary;
			} catch (Exception exception) {
				_logger.Error ("ReadFieldValuesFromDefinitions failed.", exception);
				return dictionary;
			} finally {
				ReleaseComObject (range);
			}
		}

		public bool ActivateWorkbook (Workbook workbook)
		{
			if (workbook == null) {
				return false;
			}
			try {
				workbook.Activate ();
				GetFirstVisibleWindow (workbook)?.Activate ();
				return true;
			} catch (Exception exception) {
				_logger.Error ("ActivateWorkbook failed.", exception);
				return false;
			}
		}

		public bool ActivateWorksheetByCodeName (Workbook workbook, string sheetCodeName)
		{
			Worksheet worksheet = FindWorksheetByCodeName (workbook, sheetCodeName);
			if (worksheet == null) {
				return false;
			}
			try {
				worksheet.Activate ();
				return true;
			} catch (Exception exception) {
				_logger.Error ("ActivateWorksheetByCodeName failed.", exception);
				return false;
			}
		}

		internal Workbook FindOpenWorkbook (string workbookFullName)
		{
			if (string.IsNullOrWhiteSpace (workbookFullName)) {
				return null;
			}
			try {
				string normalizedTargetPath = _pathCompatibilityService.NormalizePath (workbookFullName);
				foreach (Workbook workbook in _application.Workbooks) {
					if (IsMatchingWorkbook (workbook, normalizedTargetPath)) {
						return workbook;
					}
				}
			} catch (Exception exception) {
				_logger.Error ("FindOpenWorkbook failed.", exception);
			}
			return null;
		}

		private bool IsMatchingWorkbook (Workbook workbook, string normalizedTargetPath)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (normalizedTargetPath)) {
				return false;
			}
			string a = _pathCompatibilityService.NormalizePath (GetWorkbookFullName (workbook));
			if (string.Equals (a, normalizedTargetPath, StringComparison.OrdinalIgnoreCase)) {
				return true;
			}
			string text = _pathCompatibilityService.NormalizePath (GetWorkbookPath (workbook));
			string workbookName = GetWorkbookName (workbook);
			if (text.Length == 0 || string.IsNullOrWhiteSpace (workbookName)) {
				return false;
			}
			string a2 = _pathCompatibilityService.CombinePath (text, workbookName);
			return string.Equals (a2, normalizedTargetPath, StringComparison.OrdinalIgnoreCase);
		}

		internal IReadOnlyList<Workbook> GetOpenWorkbooks ()
		{
			List<Workbook> list = new List<Workbook> ();
			try {
				foreach (Workbook workbook in _application.Workbooks) {
					if (workbook != null) {
						list.Add (workbook);
					}
				}
			} catch (Exception exception) {
				_logger.Error ("GetOpenWorkbooks failed.", exception);
			}
			return list;
		}

		internal Workbook FindKernelWorkbook (Workbook caseWorkbook)
		{
			if (caseWorkbook == null) {
				return null;
			}
			string text = TryGetDocumentProperty (caseWorkbook, "SYSTEM_ROOT");
			if (string.IsNullOrWhiteSpace (text)) {
				return null;
			}
			string text2 = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath (text.Trim (), new PathCompatibilityService ());
			if (string.IsNullOrWhiteSpace (text2)) {
				return null;
			}
			return FindOpenWorkbook (text2);
		}

		internal Worksheet FindCaseListWorksheet (Workbook kernelWorkbook)
		{
			if (kernelWorkbook == null) {
				return null;
			}
			try {
				return kernelWorkbook.Worksheets ["案件一覧"] as Worksheet;
			} catch (Exception exception) {
				_logger.Error ("FindCaseListWorksheet failed.", exception);
				return null;
			}
		}

		internal bool TryNormalizeCaseListRowHeight (CaseContext context)
		{
			if (context == null || context.CaseListWorksheet == null || context.RegisteredRow <= 0) {
				return false;
			}
			Range range = null;
			try {
				range = context.CaseListWorksheet.Rows [context.RegisteredRow, Type.Missing] as Range;
				if (range == null) {
					return false;
				}
				range.WrapText = false;
				range.RowHeight = context.CaseListWorksheet.StandardHeight;
				_logger.Info ("CaseList row normalized. row=" + context.RegisteredRow + ", height=" + Convert.ToString (context.CaseListWorksheet.StandardHeight));
				return true;
			} catch (Exception exception) {
				_logger.Error ("TryNormalizeCaseListRowHeight failed.", exception);
				return false;
			} finally {
				ReleaseComObject (range);
			}
		}

		private string ReadValueByDefinition (Worksheet worksheet, object[,] grid, CaseListFieldDefinition definition)
		{
			if (worksheet == null || definition == null) {
				return string.Empty;
			}
			string text = TryReadNamedRangeValue (worksheet.Parent as Workbook, definition.SourceNamedRange);
			if (text.Length > 0) {
				return text;
			}
			if (!TryParseCellAddress (definition.SourceCellAddress, out var rowIndex, out var columnIndex)) {
				return string.Empty;
			}
			if (grid != null && rowIndex >= 1 && rowIndex <= grid.GetUpperBound (0) && columnIndex >= 1 && columnIndex <= grid.GetUpperBound (1)) {
				return ConvertCellValueToString (grid [rowIndex, columnIndex]);
			}
			return string.Empty;
		}

		internal Range ResolveFieldRange (Workbook workbook, Worksheet worksheet, CaseListFieldDefinition definition)
		{
			if (worksheet == null || definition == null) {
				return null;
			}
			Range range = TryResolveNamedRange (workbook, definition.SourceNamedRange);
			if (range != null) {
				return range;
			}
			if (!TryParseCellAddress (definition.SourceCellAddress, out var rowIndex, out var columnIndex)) {
				return null;
			}
			try {
				return worksheet.Cells [rowIndex, columnIndex] as Range;
			} catch (Exception exception) {
				_logger.Error ("ResolveFieldRange fallback cell resolve failed.", exception);
				return null;
			}
		}

		internal bool TryWriteFieldValue (Workbook workbook, Worksheet worksheet, CaseListFieldDefinition definition, object value)
		{
			Range range = null;
			try {
				range = ResolveFieldRange (workbook, worksheet, definition);
				if (range == null) {
					return false;
				}
				range.Value2 = value;
				return true;
			} catch (Exception exception) {
				_logger.Error ("TryWriteFieldValue failed.", exception);
				return false;
			} finally {
				ReleaseComObject (range);
			}
		}

		private string TryReadNamedRangeValue (Workbook workbook, string namedRange)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (namedRange)) {
				return string.Empty;
			}
			Name name = null;
			Range range = null;
			try {
				name = workbook.Names.Item (namedRange, Type.Missing, Type.Missing);
				range = name?.RefersToRange;
				return ConvertCellValueToString ((dynamic)range?.Value2);
			} catch {
				return string.Empty;
			} finally {
				ReleaseComObject (range);
				ReleaseComObject (name);
			}
		}

		private Range TryResolveNamedRange (Workbook workbook, string namedRange)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (namedRange)) {
				return null;
			}
			Name name = null;
			try {
				name = workbook.Names.Item (namedRange, Type.Missing, Type.Missing);
				return name?.RefersToRange;
			} catch {
				return null;
			} finally {
				ReleaseComObject (name);
			}
		}

		private static string ConvertCellValueToString (object value)
		{
			if (value == null) {
				return string.Empty;
			}
			if (value is double num) {
				return num.ToString (CultureInfo.InvariantCulture);
			}
			return Convert.ToString (value, CultureInfo.InvariantCulture) ?? string.Empty;
		}

		private static bool TryParseCellAddress (string address, out int rowIndex, out int columnIndex)
		{
			rowIndex = 0;
			columnIndex = 0;
			if (string.IsNullOrWhiteSpace (address)) {
				return false;
			}
			string text = address.Trim ().ToUpperInvariant ();
			int i;
			for (i = 0; i < text.Length && text [i] >= 'A' && text [i] <= 'Z'; i++) {
			}
			if (i == 0 || i >= text.Length) {
				return false;
			}
			for (int j = 0; j < i; j++) {
				columnIndex = columnIndex * 26 + (text [j] - 65 + 1);
			}
			string s = text.Substring (i);
			return int.TryParse (s, NumberStyles.Integer, CultureInfo.InvariantCulture, out rowIndex) && rowIndex > 0 && columnIndex > 0;
		}

		private static void ReleaseComObject (object comObject)
		{
			// Interop service が所有した COM 参照は完全解放の方針を維持する。
			ComObjectReleaseService.FinalRelease (comObject);
		}
	}
}
