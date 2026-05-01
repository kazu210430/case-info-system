using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class CaseListRegistrationService
	{
		private static readonly string[] AttorneyPrefixesToTrim = new string[3] { "弁護士 ", "弁護士\u3000", "弁護士" };

		private const string HomeSheetCodeName = "shHOME";

		private const int RegisteredDateColumn = 1;

		private const string CaseListRegisteredPropName = "CASELIST_REGISTERED";

		private const string CaseListRowPropName = "CASELIST_ROW";

		private const string SuppressUiOnOpenPropName = "SUPPRESS_UI_ON_OPEN";

		private const string SuppressHomeOnActivatePropName = "SUPPRESS_VSTO_HOME_ON_ACTIVATE";

		private const string TaskPaneSnapshotCacheCountPropName = "TASKPANE_SNAPSHOT_CACHE_COUNT";

		private readonly ExcelInteropService _excelInteropService;

		private readonly KernelWorkbookResolverService _kernelWorkbookResolverService;

		private readonly CaseDataSnapshotFactory _caseDataSnapshotFactory;

		private readonly CaseListFieldDefinitionRepository _fieldDefinitionRepository;

		private readonly CaseListHeaderRepository _headerRepository;

		private readonly CaseListMappingRepository _mappingRepository;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly TaskPaneSnapshotCacheService _taskPaneSnapshotCacheService;

		private readonly Logger _logger;

		internal CaseListRegistrationService (ExcelInteropService excelInteropService, KernelWorkbookResolverService kernelWorkbookResolverService, CaseDataSnapshotFactory caseDataSnapshotFactory, CaseListFieldDefinitionRepository fieldDefinitionRepository, CaseListHeaderRepository headerRepository, CaseListMappingRepository mappingRepository, AccountingWorkbookService accountingWorkbookService, TaskPaneSnapshotCacheService taskPaneSnapshotCacheService, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_kernelWorkbookResolverService = kernelWorkbookResolverService ?? throw new ArgumentNullException ("kernelWorkbookResolverService");
			_caseDataSnapshotFactory = caseDataSnapshotFactory ?? throw new ArgumentNullException ("caseDataSnapshotFactory");
			_fieldDefinitionRepository = fieldDefinitionRepository ?? throw new ArgumentNullException ("fieldDefinitionRepository");
			_headerRepository = headerRepository ?? throw new ArgumentNullException ("headerRepository");
			_mappingRepository = mappingRepository ?? throw new ArgumentNullException ("mappingRepository");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_taskPaneSnapshotCacheService = taskPaneSnapshotCacheService ?? throw new ArgumentNullException ("taskPaneSnapshotCacheService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal CaseListRegistrationResult Execute (Workbook caseWorkbook)
		{
			if (caseWorkbook == null) {
				throw new ArgumentNullException ("caseWorkbook");
			}
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (caseWorkbook, "shHOME");
			if (worksheet == null) {
				return new CaseListRegistrationResult {
					Success = false,
					Message = "案件情報シートを取得できません。"
				};
			}
			bool openedNow;
			Workbook workbook = _kernelWorkbookResolverService.ResolveOrOpen (caseWorkbook, out openedNow);
			if (workbook == null) {
				return new CaseListRegistrationResult {
					Success = false,
					Message = "案件情報System_Kernel を開けませんでした。"
				};
			}
			IReadOnlyDictionary<string, string> sourceValues = ReadHomeValues (caseWorkbook, workbook, worksheet);
			Worksheet worksheet2 = _excelInteropService.FindCaseListWorksheet (workbook);
			if (worksheet2 == null) {
				return new CaseListRegistrationResult {
					Success = false,
					Message = "Kernelブックにシート「案件一覧」が見つかりません。"
				};
			}
			IReadOnlyDictionary<string, CaseListFieldDefinition> fieldDefinitions = _fieldDefinitionRepository.LoadDefinitions (workbook);
			IReadOnlyList<CaseListHeaderDefinition> headerDefinitions = _headerRepository.LoadDefinitions (workbook);
			IReadOnlyList<CaseListMappingDefinition> mappingDefinitions = _mappingRepository.LoadEnabledDefinitions (workbook);
			string text = ValidateDefinitions (fieldDefinitions, headerDefinitions, mappingDefinitions);
			if (!string.IsNullOrWhiteSpace (text)) {
				return new CaseListRegistrationResult {
					Success = false,
					Message = text
				};
			}
			int nextCaseListRow = GetNextCaseListRow (worksheet2);
			WriteCaseListRow (worksheet2, nextCaseListRow, sourceValues, fieldDefinitions, headerDefinitions, mappingDefinitions);
			_excelInteropService.SetDocumentProperty (workbook, "SUPPRESS_UI_ON_OPEN", "0");
			_excelInteropService.SetDocumentProperty (workbook, "SUPPRESS_VSTO_HOME_ON_ACTIVATE", "1");
			_excelInteropService.SetDocumentProperty (caseWorkbook, "CASELIST_REGISTERED", "1");
			_excelInteropService.SetDocumentProperty (caseWorkbook, "CASELIST_ROW", nextCaseListRow.ToString ());
			_excelInteropService.SetDocumentProperty (caseWorkbook, "TASKPANE_SNAPSHOT_CACHE_COUNT", "0");
			_taskPaneSnapshotCacheService.ClearCaseSnapshotCacheChunks (caseWorkbook);
			_logger.Info ("Case workbook task pane snapshot cache invalidated after case-list registration. caseListRegistered=1, caseListRow=" + nextCaseListRow);
			if (!openedNow) {
				_logger.Info ("Case list registration reused open kernel workbook.");
			}
			return new CaseListRegistrationResult {
				Success = true,
				RegisteredRow = nextCaseListRow,
				Message = "案件一覧登録が完了しました。（案件一覧 行: " + nextCaseListRow + "）"
			};
		}

		private IReadOnlyDictionary<string, string> ReadHomeValues (Workbook caseWorkbook, Workbook kernelWorkbook, Worksheet homeWorksheet)
		{
			if (caseWorkbook == null) {
				throw new ArgumentNullException ("caseWorkbook");
			}
			if (kernelWorkbook == null) {
				throw new ArgumentNullException ("kernelWorkbook");
			}
			if (homeWorksheet == null) {
				throw new ArgumentNullException ("homeWorksheet");
			}
			IReadOnlyDictionary<string, string> readOnlyDictionary = _caseDataSnapshotFactory.Create (caseWorkbook, kernelWorkbook)?.Values;
			if (readOnlyDictionary == null || readOnlyDictionary.Count == 0) {
				throw new InvalidOperationException ("案件情報シートを管理定義に従って読み取れませんでした。");
			}
			return readOnlyDictionary;
		}

		private static int GetNextCaseListRow (Worksheet caseListWorksheet)
		{
			try {
				if (caseListWorksheet.ListObjects.Count > 0) {
					ListObject listObject = caseListWorksheet.ListObjects [1];
					int num = TryGetReusableEmptyTableRow (listObject);
					if (num > 0) {
						return num;
					}
					ListRow listRow = listObject.ListRows.Add (Type.Missing);
					try {
						return listRow.Range.Row;
					} finally {
						ReleaseComObject (listRow);
					}
				}
			} catch {
			}
			int num2 = ((dynamic)caseListWorksheet.Cells [caseListWorksheet.Rows.Count, 1]).End [XlDirection.xlUp].Row;
			return (num2 < 3) ? 3 : (num2 + 1);
		}

		private static int TryGetReusableEmptyTableRow (ListObject listObject)
		{
			if (listObject == null) {
				throw new ArgumentNullException ("listObject");
			}
			Range range = null;
			try {
				range = listObject.DataBodyRange;
				if (range == null) {
					return 0;
				}
				int count = range.Rows.Count;
				int count2 = range.Columns.Count;
				if (count <= 0 || count2 <= 0) {
					return 0;
				}
				if (!(range.Value2 is object[,] values)) {
					return 0;
				}
				for (int num = count; num >= 1; num--) {
					if (IsEmptyTableRow (values, num, count2)) {
						return range.Row + num - 1;
					}
				}
				return 0;
			} finally {
				ReleaseComObject (range);
			}
		}

		private static bool IsEmptyTableRow (object[,] values, int rowIndex, int columnCount)
		{
			if (values == null || rowIndex <= 0 || columnCount <= 0) {
				return false;
			}
			for (int i = 1; i <= columnCount; i++) {
				object obj = values [rowIndex, i];
				if (obj != null && !string.IsNullOrWhiteSpace (Convert.ToString (obj))) {
					return false;
				}
			}
			return true;
		}

		private void WriteCaseListRow (Worksheet caseListWorksheet, int nextRow, IReadOnlyDictionary<string, string> sourceValues, IReadOnlyDictionary<string, CaseListFieldDefinition> fieldDefinitions, IReadOnlyList<CaseListHeaderDefinition> headerDefinitions, IReadOnlyList<CaseListMappingDefinition> mappingDefinitions)
		{
			if (caseListWorksheet == null) {
				throw new ArgumentNullException ("caseListWorksheet");
			}
			if (sourceValues == null) {
				throw new ArgumentNullException ("sourceValues");
			}
			_accountingWorkbookService.WriteCellValue (caseListWorksheet, BuildCellAddress (nextRow, RegisteredDateColumn), DateTime.Today);
			IReadOnlyDictionary<string, int> readOnlyDictionary = ReadActualCaseListHeaderMap (caseListWorksheet);
			IReadOnlyDictionary<string, int> readOnlyDictionary2 = BuildManagedHeaderMap (headerDefinitions);
			int num = 0;
			foreach (CaseListMappingDefinition mappingDefinition in mappingDefinitions) {
				if (string.Equals (mappingDefinition.MappingType, "Direct", StringComparison.OrdinalIgnoreCase) && sourceValues.TryGetValue (mappingDefinition.SourceFieldKey, out var value) && fieldDefinitions.TryGetValue (mappingDefinition.SourceFieldKey, out var value2) && readOnlyDictionary2.TryGetValue (mappingDefinition.TargetHeaderName, out var value3)) {
					if (!readOnlyDictionary.TryGetValue (mappingDefinition.TargetHeaderName, out var value4)) {
						throw new InvalidOperationException ("案件一覧シートに管理定義ヘッダが存在しません。 header=" + mappingDefinition.TargetHeaderName);
					}
					if (value3 != value4) {
						throw new InvalidOperationException ("案件一覧シートの列配置が管理定義と一致しません。 header=" + mappingDefinition.TargetHeaderName + ", managedColumn=" + value3 + ", actualColumn=" + value4);
					}
					_accountingWorkbookService.WriteCellValue (caseListWorksheet, BuildCellAddress (nextRow, value4), NormalizeForCaseList (value2, mappingDefinition, value));
					num++;
				}
			}
			_logger.Info ("Case list row written by managed mapping. row=" + nextRow + ", writtenColumns=" + num);
		}

		private static string ValidateDefinitions (IReadOnlyDictionary<string, CaseListFieldDefinition> fieldDefinitions, IReadOnlyList<CaseListHeaderDefinition> headerDefinitions, IReadOnlyList<CaseListMappingDefinition> mappingDefinitions)
		{
			if (fieldDefinitions == null || fieldDefinitions.Count == 0) {
				return "Kernelブックの管理シート CaseList_FieldInventory を読み取れません。";
			}
			if (headerDefinitions == null || headerDefinitions.Count == 0) {
				return "Kernelブックの管理シート CaseList_Headers を読み取れません。";
			}
			if (mappingDefinitions == null || mappingDefinitions.Count == 0) {
				return "Kernelブックの管理シート CaseList_Mapping を読み取れません。";
			}
			HashSet<string> hashSet = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
			foreach (CaseListHeaderDefinition headerDefinition in headerDefinitions) {
				if (!string.IsNullOrWhiteSpace (headerDefinition.HeaderName)) {
					hashSet.Add (headerDefinition.HeaderName);
				}
			}
			foreach (CaseListMappingDefinition mappingDefinition in mappingDefinitions) {
				if (!fieldDefinitions.ContainsKey (mappingDefinition.SourceFieldKey)) {
					return "CaseList_Mapping に FieldInventory 未定義の項目があります。 sourceFieldKey=" + mappingDefinition.SourceFieldKey;
				}
				if (!hashSet.Contains (mappingDefinition.TargetHeaderName)) {
					return "CaseList_Mapping に Headers 未定義のヘッダがあります。 targetHeaderName=" + mappingDefinition.TargetHeaderName;
				}
			}
			return string.Empty;
		}

		private static IReadOnlyDictionary<string, int> BuildManagedHeaderMap (IReadOnlyList<CaseListHeaderDefinition> headerDefinitions)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
			if (headerDefinitions == null) {
				return dictionary;
			}
			foreach (CaseListHeaderDefinition headerDefinition in headerDefinitions) {
				int num = ConvertColumnAddressToIndex ((headerDefinition == null) ? string.Empty : headerDefinition.CellAddress);
				string text = ((headerDefinition == null) ? string.Empty : (headerDefinition.HeaderName ?? string.Empty).Trim ());
				if (num > 0 && text.Length != 0 && !dictionary.ContainsKey (text)) {
					dictionary.Add (text, num);
				}
			}
			return dictionary;
		}

		private static IReadOnlyDictionary<string, int> ReadActualCaseListHeaderMap (Worksheet caseListWorksheet)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
			Range range = null;
			try {
				int num = ((dynamic)caseListWorksheet.Cells [2, caseListWorksheet.Columns.Count]).End [XlDirection.xlToLeft].Column;
				if (num < 1) {
					return dictionary;
				}
				range = caseListWorksheet.Range [(dynamic)caseListWorksheet.Cells [2, 1], (dynamic)caseListWorksheet.Cells [2, num]];
				if (!(range.Value2 is object[,] array)) {
					return dictionary;
				}
				int upperBound = array.GetUpperBound (1);
				for (int i = 1; i <= upperBound; i++) {
					string text = (Convert.ToString (array [1, i]) ?? string.Empty).Trim ();
					if (text.Length != 0 && !dictionary.ContainsKey (text)) {
						dictionary.Add (text, i);
					}
				}
				return dictionary;
			} finally {
				ReleaseComObject (range);
			}
		}

		private static object NormalizeForCaseList (CaseListFieldDefinition fieldDefinition, CaseListMappingDefinition mapping, string value)
		{
			if (string.IsNullOrEmpty (value)) {
				return string.Empty;
			}
			string a = ResolveNormalizeRule (fieldDefinition, mapping);
			string a2 = ResolveDataType (fieldDefinition, mapping);
			string text = value.Trim ();
			if (string.Equals (a, "TrimHalfWidth", StringComparison.OrdinalIgnoreCase)) {
				text = ConvertFullWidthAsciiToHalfWidth (text).Trim ();
			} else if (string.Equals (a, "Trim", StringComparison.OrdinalIgnoreCase)) {
				text = text.Trim ();
			} else if (string.Equals (a, "StripAttorneyPrefix", StringComparison.OrdinalIgnoreCase)) {
				text = StripAttorneyPrefix (text);
			}
			if (text.Length == 0) {
				return string.Empty;
			}
			if (string.Equals (a2, "Date", StringComparison.OrdinalIgnoreCase) && DateTime.TryParse (text, out var result)) {
				return result;
			}
			if (string.Equals (a, "LowerInvariant", StringComparison.OrdinalIgnoreCase) || (fieldDefinition != null && (fieldDefinition.FieldKey ?? string.Empty).IndexOf ("mail", StringComparison.OrdinalIgnoreCase) >= 0)) {
				return text.ToLowerInvariant ();
			}
			return text;
		}

		private static string StripAttorneyPrefix (string value)
		{
			if (string.IsNullOrWhiteSpace (value)) {
				return string.Empty;
			}
			string text = value.Trim ();
			string[] attorneyPrefixesToTrim = AttorneyPrefixesToTrim;
			foreach (string text2 in attorneyPrefixesToTrim) {
				if (text.StartsWith (text2, StringComparison.Ordinal)) {
					return text.Substring (text2.Length).Trim ();
				}
			}
			return text;
		}

		private static string ResolveNormalizeRule (CaseListFieldDefinition fieldDefinition, CaseListMappingDefinition mapping)
		{
			string text = ((mapping == null) ? string.Empty : (mapping.NormalizeRule ?? string.Empty));
			if (!string.IsNullOrWhiteSpace (text)) {
				return text.Trim ();
			}
			return (fieldDefinition == null) ? string.Empty : (fieldDefinition.NormalizeRule ?? string.Empty).Trim ();
		}

		private static string ResolveDataType (CaseListFieldDefinition fieldDefinition, CaseListMappingDefinition mapping)
		{
			string text = ((mapping == null) ? string.Empty : (mapping.DataType ?? string.Empty));
			if (!string.IsNullOrWhiteSpace (text)) {
				return text.Trim ();
			}
			return (fieldDefinition == null) ? string.Empty : (fieldDefinition.DataType ?? string.Empty).Trim ();
		}

		private static int ConvertColumnAddressToIndex (string cellAddress)
		{
			if (string.IsNullOrWhiteSpace (cellAddress)) {
				return 0;
			}
			int num = 0;
			string text = cellAddress.Trim ().ToUpperInvariant ();
			foreach (char c in text) {
				if (c < 'A' || c > 'Z') {
					break;
				}
				num = num * 26 + (c - 65 + 1);
			}
			return num;
		}

		private static string BuildCellAddress (int rowIndex, int columnIndex)
		{
			if (rowIndex <= 0 || columnIndex <= 0) {
				return string.Empty;
			}
			return ConvertColumnIndexToAddress (columnIndex) + rowIndex.ToString ();
		}

		private static string ConvertColumnIndexToAddress (int columnIndex)
		{
			if (columnIndex <= 0) {
				return string.Empty;
			}
			string text = string.Empty;
			int num = columnIndex;
			while (num > 0) {
				num--;
				text = (char)('A' + num % 26) + text;
				num /= 26;
			}
			return text;
		}

		private static string ConvertFullWidthAsciiToHalfWidth (string value)
		{
			char[] array = new char[value.Length];
			for (int i = 0; i < value.Length; i++) {
				char c = value [i];
				if (c == '\u3000') {
					array [i] = ' ';
				} else if (c >= '！' && c <= '～') {
					array [i] = (char)(c - 65248);
				} else {
					array [i] = c;
				}
			}
			return new string (array);
		}

		private static void ReleaseComObject (object comObject)
		{
			// 登録処理で所有した COM 参照は完全解放の方針を維持する。
			CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease (comObject);
		}
	}
}
