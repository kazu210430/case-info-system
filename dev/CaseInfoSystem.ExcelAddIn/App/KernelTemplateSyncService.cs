using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelTemplateSyncService
	{
		private sealed class MasterSheetSnapshot
		{
			private readonly Dictionary<string, long> _tabBackColors;

			internal List<MasterSheetRow> Rows { get; }

			internal MasterSheetSnapshot (List<MasterSheetRow> rows, Dictionary<string, long> tabBackColors)
			{
				Rows = rows ?? new List<MasterSheetRow> ();
				_tabBackColors = tabBackColors ?? new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase);
			}

			internal long GetTabBackColor (string tabName)
			{
				if (string.IsNullOrWhiteSpace (tabName)) {
					return 0L;
				}
				long value;
				return _tabBackColors.TryGetValue (tabName, out value) ? value : 0;
			}
		}

		private sealed class MasterSheetRow
		{
			internal string Key { get; }

			internal string TemplateFileName { get; }

			internal string Caption { get; }

			internal string TabName { get; }

			internal long FillColor { get; }

			internal MasterSheetRow (string key, string templateFileName, string caption, string tabName, long fillColor)
			{
				Key = key ?? string.Empty;
				TemplateFileName = templateFileName ?? string.Empty;
				Caption = caption ?? string.Empty;
				TabName = tabName ?? string.Empty;
				FillColor = fillColor;
			}
		}

		private sealed class AppState
		{
			internal bool ScreenUpdating { get; set; }

			internal bool EnableEvents { get; set; }

			internal bool DisplayAlerts { get; set; }
		}

		private sealed class SheetProtectionState
		{
			internal bool IsProtected { get; set; }

			internal bool AllowFormattingCells { get; set; }

			internal bool AllowFormattingColumns { get; set; }

			internal bool AllowFormattingRows { get; set; }

			internal bool AllowInsertingColumns { get; set; }

			internal bool AllowInsertingRows { get; set; }

			internal bool AllowInsertingHyperlinks { get; set; }

			internal bool AllowDeletingColumns { get; set; }

			internal bool AllowDeletingRows { get; set; }

			internal bool AllowSorting { get; set; }

			internal bool AllowFiltering { get; set; }

			internal bool AllowUsingPivotTables { get; set; }
		}

		private const string TemplateFolderName = "雛形";

		private const string MasterSheetCodeName = "shMasterList";

		private const string MasterSheetName = "雛形一覧";

		private const int MasterListFirstDataRow = 3;

		private const int MasterListMaxKeyCount = 99;

		private const int ColumnA = 1;

		private const int ColumnB = 2;

		private const int ColumnC = 3;

		private const int ColumnD = 4;

		private const int ColumnE = 5;

		private const int ColumnF = 6;

		private const string TaskPaneMasterVersionProp = "TASKPANE_MASTER_VERSION";

		private const string TaskPaneBaseCacheCountProp = "TASKPANE_BASE_SNAPSHOT_COUNT";

		private const string TaskPaneBaseCachePartPropPrefix = "TASKPANE_BASE_SNAPSHOT_";

		private const string TaskPaneBaseMasterVersionProp = "TASKPANE_BASE_MASTER_VERSION";

		private const int TaskPaneCacheChunkSize = 240;

		private const string AllTabCaption = "全て";

		private const string DefaultTabCaption = "その他";

		private const string CaseListActionCaption = "案件一覧登録（未了）";

		private const string AccountingActionCaption = "会計書類セット";

		private readonly Application _application;

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly ExcelInteropService _excelInteropService;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly CaseListFieldDefinitionRepository _caseListFieldDefinitionRepository;

		private readonly WordTemplateRegistrationValidationService _wordTemplateRegistrationValidationService;

		private readonly MasterTemplateCatalogService _masterTemplateCatalogService;

		private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

		private readonly Logger _logger;

		internal KernelTemplateSyncService (Application application, KernelWorkbookService kernelWorkbookService, ExcelInteropService excelInteropService, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository, WordTemplateRegistrationValidationService wordTemplateRegistrationValidationService, MasterTemplateCatalogService masterTemplateCatalogService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException ("caseListFieldDefinitionRepository");
			_wordTemplateRegistrationValidationService = wordTemplateRegistrationValidationService ?? throw new ArgumentNullException ("wordTemplateRegistrationValidationService");
			_masterTemplateCatalogService = masterTemplateCatalogService ?? throw new ArgumentNullException ("masterTemplateCatalogService");
			_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal KernelTemplateSyncResult Execute ()
		{
			Workbook openKernelWorkbook = _kernelWorkbookService.GetOpenKernelWorkbook ();
			if (openKernelWorkbook == null) {
				throw new InvalidOperationException ("Kernel ブックを開いてから実行してください。");
			}
			Stopwatch stopwatch = Stopwatch.StartNew ();
			AppState state = SaveAppState ();
			Worksheet worksheet = null;
			SheetProtectionState sheetProtectionState = null;
			try {
				ApplyFastSettings ();
				worksheet = GetMasterListSheet (openKernelWorkbook);
				sheetProtectionState = SaveSheetProtectionState (worksheet);
				if (sheetProtectionState.IsProtected) {
					worksheet.Unprotect (string.Empty);
				}
				ValidateMasterListSheet (worksheet);
				string text = ResolveTemplateDirectory (openKernelWorkbook);
				IReadOnlyCollection<string> readOnlyCollection = LoadDefinedTemplateTags (openKernelWorkbook);
				if (readOnlyCollection.Count == 0) {
					return new KernelTemplateSyncResult {
						Success = false,
						TemplateDirectory = text,
						Message = "Kernelブックの管理シート CaseList_FieldInventory を読み取れません。"
					};
				}
				TemplateRegistrationValidationSummary templateRegistrationValidationSummary = _wordTemplateRegistrationValidationService.Validate (text, readOnlyCollection);
				if (templateRegistrationValidationSummary.DetectedFileCount == 0) {
					return new KernelTemplateSyncResult {
						Success = false,
						TemplateDirectory = text,
						DetectedCount = 0,
						TemplateResults = templateRegistrationValidationSummary.TemplateResults,
						Message = "雛形フォルダに Word 雛形 (.docx / .dotx / .docm / .dotm) が見つかりませんでした。" + Environment.NewLine + "フォルダ: " + text
					};
				}
				IReadOnlyList<TemplateRegistrationValidationEntry> validTemplates = templateRegistrationValidationSummary.GetValidTemplates ();
				int updatedCount = WriteToMasterList (worksheet, validTemplates);
				int masterVersion = IncrementTaskPaneMasterVersion (openKernelWorkbook);
				openKernelWorkbook.Save ();
				string snapshotText = BuildTaskPaneSnapshot (worksheet, openKernelWorkbook, masterVersion);
				string errorMessage;
				bool baseSyncSucceeded = TrySyncTaskPaneSnapshotToBase (openKernelWorkbook, snapshotText, masterVersion, out errorMessage);
				_masterTemplateCatalogService.InvalidateCache ();
				_logger.Info ("Kernel template sync completed. updatedCount=" + updatedCount + ", detectedCount=" + templateRegistrationValidationSummary.DetectedFileCount + ", excludedCount=" + templateRegistrationValidationSummary.ExcludedTemplateCount + ", warningCount=" + templateRegistrationValidationSummary.WarningFileCount + ", masterVersion=" + masterVersion);
				return new KernelTemplateSyncResult {
					Success = true,
					UpdatedCount = updatedCount,
					DetectedCount = templateRegistrationValidationSummary.DetectedFileCount,
					ExcludedCount = templateRegistrationValidationSummary.ExcludedTemplateCount,
					WarningCount = templateRegistrationValidationSummary.WarningFileCount,
					MasterVersion = masterVersion,
					TemplateDirectory = text,
					TemplateResults = templateRegistrationValidationSummary.TemplateResults,
					BaseSyncError = errorMessage,
					Message = BuildCompletedMessage (text, updatedCount, templateRegistrationValidationSummary, stopwatch.Elapsed, masterVersion, baseSyncSucceeded, errorMessage)
				};
			} finally {
				if (worksheet != null && sheetProtectionState != null) {
					RestoreSheetProtectionState (worksheet, sheetProtectionState);
				}
				RestoreAppState (state);
			}
		}

		private IReadOnlyCollection<string> LoadDefinedTemplateTags (Workbook kernelWorkbook)
		{
			IReadOnlyDictionary<string, CaseListFieldDefinition> readOnlyDictionary = _caseListFieldDefinitionRepository.LoadDefinitions (kernelWorkbook);
			if (readOnlyDictionary == null || readOnlyDictionary.Count == 0) {
				return Array.Empty<string> ();
			}
			return readOnlyDictionary.Keys
				.Where (key => !string.IsNullOrWhiteSpace (key))
				.Select (key => key.Trim ())
				.ToArray ();
		}

		private int WriteToMasterList (Worksheet masterSheet, IReadOnlyList<TemplateRegistrationValidationEntry> templates)
		{
			int num = 3;
			int num2 = 101;
			masterSheet.Range [(dynamic)masterSheet.Cells [num, 1], (dynamic)masterSheet.Cells [num2, 3]].ClearContents ();
			object[,] array = new object[99, 3];
			int num3 = 0;
			if (templates == null) {
				return 0;
			}
			foreach (TemplateRegistrationValidationEntry template in templates) {
				if (template == null || !int.TryParse (template.Key, out var result)) {
					continue;
				}
				int num4 = result + 3 - 1;
				if (num4 >= num && num4 <= num2) {
					string text = template.FileName ?? string.Empty;
					int num5 = num4 - num;
					array [num5, 0] = result.ToString ("00");
					array [num5, 1] = text;
					array [num5, 2] = template.DisplayName ?? ExtractDocumentName (text);
					num3++;
				}
			}
			_accountingWorkbookService.WriteRangeValues (masterSheet, "$A$" + num.ToString () + ":$C$" + num2.ToString (), array);
			return num3;
		}

		private int IncrementTaskPaneMasterVersion (Workbook kernelWorkbook)
		{
			string s = _excelInteropService.TryGetDocumentProperty (kernelWorkbook, "TASKPANE_MASTER_VERSION");
			int num = 1;
			if (int.TryParse (s, out var result)) {
				num = result + 1;
			}
			if (num < 1) {
				num = 1;
			}
			_excelInteropService.SetDocumentProperty (kernelWorkbook, "TASKPANE_MASTER_VERSION", num.ToString ());
			return num;
		}

		private bool TrySyncTaskPaneSnapshotToBase (Workbook kernelWorkbook, string snapshotText, int masterVersion, out string errorMessage)
		{
			errorMessage = string.Empty;
			string systemRoot = ResolveSystemRoot (kernelWorkbook);
			string text = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath (systemRoot, _pathCompatibilityService);
			if (!_pathCompatibilityService.FileExistsSafe (text)) {
				errorMessage = "Base が見つかりません: " + text;
				return false;
			}
			Workbook workbook = _excelInteropService.FindOpenWorkbook (text);
			bool flag = workbook != null;
			try {
				if (workbook == null) {
					workbook = _application.Workbooks.Open (text, 0, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				}
				SaveSnapshotToBaseWorkbook (workbook, snapshotText, masterVersion);
				workbook.Save ();
				return true;
			} catch (Exception ex) {
				errorMessage = ex.Message;
				_logger.Error ("Kernel template sync base snapshot failed.", ex);
				return false;
			} finally {
				if (!flag && workbook != null) {
					using (_caseWorkbookLifecycleService.BeginManagedCloseScope (workbook)) {
						workbook.Close (false, Type.Missing, Type.Missing);
					}
				}
			}
		}

		private void SaveSnapshotToBaseWorkbook (Workbook baseWorkbook, string snapshotText, int masterVersion)
		{
			string s = _excelInteropService.TryGetDocumentProperty (baseWorkbook, "TASKPANE_BASE_SNAPSHOT_COUNT");
			int result;
			int num = (int.TryParse (s, out result) ? result : 0);
			if (string.IsNullOrEmpty (snapshotText)) {
				_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_BASE_SNAPSHOT_COUNT", "0");
				_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_BASE_MASTER_VERSION", masterVersion.ToString ());
				return;
			}
			int num2 = (snapshotText.Length - 1) / 240 + 1;
			_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_BASE_SNAPSHOT_COUNT", num2.ToString ());
			_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_BASE_MASTER_VERSION", masterVersion.ToString ());
			_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_MASTER_VERSION", masterVersion.ToString ());
			for (int i = 1; i <= num2; i++) {
				int num3 = (i - 1) * 240;
				int length = Math.Min (240, snapshotText.Length - num3);
				_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_BASE_SNAPSHOT_" + i.ToString ("00"), snapshotText.Substring (num3, length));
			}
			for (int j = num2 + 1; j <= num; j++) {
				_excelInteropService.SetDocumentProperty (baseWorkbook, "TASKPANE_BASE_SNAPSHOT_" + j.ToString ("00"), string.Empty);
			}
		}

		private string BuildTaskPaneSnapshot (Worksheet masterSheet, Workbook kernelWorkbook, int masterVersion)
		{
			List<string> list = new List<string> ();
			string systemRoot = ResolveSystemRoot (kernelWorkbook);
			string text = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath (systemRoot, _pathCompatibilityService);
			string text2 = Path.GetFileName (text);
			if (string.IsNullOrWhiteSpace (text2)) {
				text2 = WorkbookFileNameResolver.BuildBaseWorkbookName (Path.GetExtension (text));
			}
			list.Add (JoinFields ("META", "2", text2, text, BuildPreferredPaneWidth (masterSheet).ToString (), masterVersion.ToString ()));
			list.Add (JoinFields ("SPECIAL", "btnCaseList", "案件一覧登録（未了）", "caselist", string.Empty, "18", "16", "128", "32", ColorToString (248, 225, 193)));
			list.Add (JoinFields ("SPECIAL", "btnAccounting", "会計書類セット", "accounting", string.Empty, "18", "64", "128", "32", ColorToString (226, 239, 218)));
			Dictionary<string, int> dictionary = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
			MasterSheetSnapshot masterSheetSnapshot = ReadMasterSheetSnapshot (masterSheet);
			Dictionary<string, int> dictionary2 = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
			for (int i = 0; i < masterSheetSnapshot.Rows.Count; i++) {
				MasterSheetRow masterSheetRow = masterSheetSnapshot.Rows [i];
				string key = masterSheetRow.Key;
				string templateFileName = masterSheetRow.TemplateFileName;
				string caption = masterSheetRow.Caption;
				string tabName = masterSheetRow.TabName;
				if (key.Length != 0 && caption.Length != 0) {
					if (!dictionary.ContainsKey (tabName)) {
						int value = dictionary.Count + 1;
						dictionary.Add (tabName, value);
						dictionary2 [tabName] = 0;
						long tabBackColor = masterSheetSnapshot.GetTabBackColor (tabName);
						list.Add (JoinFields ("TAB", value.ToString (), tabName, tabBackColor.ToString ()));
					}
					int num = (dictionary2 [tabName] = ((!dictionary2.ContainsKey (tabName)) ? 1 : (dictionary2 [tabName] + 1)));
					list.Add (JoinFields ("DOC", "btnDoc_" + key, key, caption, "doc", tabName, num.ToString (), masterSheetRow.FillColor.ToString (), templateFileName));
				}
			}
			if (!dictionary.ContainsKey ("全て")) {
				int value2 = dictionary.Count + 1;
				dictionary.Add ("全て", value2);
				list.Add (JoinFields ("TAB", value2.ToString (), "全て", ColorTranslator.ToOle (Color.FromArgb (255, 255, 255)).ToString ()));
			}
			return string.Join (Environment.NewLine, list.ToArray ());
		}

		private static string BuildCompletedMessage (string templateDirectory, int updatedCount, TemplateRegistrationValidationSummary validationSummary, TimeSpan elapsed, int masterVersion, bool baseSyncSucceeded, string baseSyncError)
		{
			StringBuilder stringBuilder = new StringBuilder ();
			List<TemplateRegistrationValidationEntry> list = ((validationSummary == null) ? new List<TemplateRegistrationValidationEntry> () : validationSummary.TemplateResults.Where (entry => entry != null && !entry.IsValid).OrderBy (entry => entry.FileName ?? string.Empty, StringComparer.OrdinalIgnoreCase).ToList ());
			List<TemplateRegistrationValidationEntry> list2 = ((validationSummary == null) ? new List<TemplateRegistrationValidationEntry> () : validationSummary.TemplateResults.Where (entry => entry != null && entry.HasWarnings).OrderBy (entry => entry.FileName ?? string.Empty, StringComparer.OrdinalIgnoreCase).ToList ());
			stringBuilder.Append ("雛形登録・更新が完了しました。").AppendLine ().Append ("登録成功: ")
				.Append (updatedCount.ToString ())
				.AppendLine ()
				.Append ("登録除外: ")
				.Append (list.Count.ToString ())
				.AppendLine ()
				.Append ("警告: ")
				.Append (list2.Count.ToString ())
				.AppendLine ()
				.Append ("フォルダ: ")
				.Append (templateDirectory)
				.AppendLine ()
				.Append ("検出件数: ")
				.Append ((validationSummary == null) ? "0" : validationSummary.DetectedFileCount.ToString ())
				.AppendLine ()
				.Append ("処理時間(秒): ")
				.Append (elapsed.TotalSeconds.ToString ("0.00"))
				.AppendLine ()
				.Append ("Master版: ")
				.Append (masterVersion.ToString ());
			if (list.Count > 0) {
				stringBuilder.AppendLine ().AppendLine ().Append ("登録除外:");
				foreach (TemplateRegistrationValidationEntry item in list) {
					stringBuilder.AppendLine ().Append ("- ").Append (item.FileName ?? string.Empty);
					foreach (string error in item.Errors) {
						stringBuilder.AppendLine ().Append ("  - ").Append (error ?? string.Empty);
					}
				}
			}
			if (list2.Count > 0) {
				stringBuilder.AppendLine ().AppendLine ().Append ("警告:");
				foreach (TemplateRegistrationValidationEntry item2 in list2) {
					stringBuilder.AppendLine ().Append ("- ").Append (item2.FileName ?? string.Empty);
					foreach (string warning in item2.Warnings) {
						stringBuilder.AppendLine ().Append ("  - ").Append (warning ?? string.Empty);
					}
				}
			}
			if (!baseSyncSucceeded) {
				stringBuilder.AppendLine ().AppendLine ().Append ("【注意】Base への初期 Task Pane 定義反映に失敗しました。")
					.AppendLine ()
					.Append (baseSyncError ?? string.Empty);
			}
			return stringBuilder.ToString ();
		}

		private string ResolveTemplateDirectory (Workbook kernelWorkbook)
		{
			string left = ResolveSystemRoot (kernelWorkbook);
			string text = _pathCompatibilityService.CombinePath (left, "雛形");
			if (!Directory.Exists (text)) {
				throw new InvalidOperationException ("雛形フォルダが見つかりません: " + text);
			}
			return text;
		}

		private string ResolveSystemRoot (Workbook kernelWorkbook)
		{
			string text = _pathCompatibilityService.NormalizePath (_excelInteropService.TryGetDocumentProperty (kernelWorkbook, "SYSTEM_ROOT"));
			if (!string.IsNullOrWhiteSpace (text)) {
				return text;
			}
			text = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookPath (kernelWorkbook));
			if (string.IsNullOrWhiteSpace (text)) {
				throw new InvalidOperationException ("SYSTEM_ROOT を解決できませんでした。");
			}
			return text;
		}

		private Worksheet GetMasterListSheet (Workbook kernelWorkbook)
		{
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (kernelWorkbook, "shMasterList");
			if (worksheet != null) {
				return worksheet;
			}
			try {
				worksheet = kernelWorkbook.Worksheets ["雛形一覧"] as Worksheet;
			} catch {
				worksheet = null;
			}
			if (worksheet == null) {
				throw new InvalidOperationException ("雛形登録シートが見つかりません。");
			}
			return worksheet;
		}

		private static void ValidateMasterListSheet (Worksheet worksheet)
		{
			if (worksheet == null) {
				throw new ArgumentNullException ("worksheet");
			}
			if (worksheet.ProtectContents) {
				throw new InvalidOperationException ("雛形登録シートが保護されています。保護解除してから実行してください。");
			}
		}

		private static string ExtractDocumentName (string fileName)
		{
			if (string.IsNullOrWhiteSpace (fileName)) {
				return string.Empty;
			}
			string text = Path.GetFileNameWithoutExtension (fileName) ?? string.Empty;
			return (text.Length >= 4) ? text.Substring (3) : string.Empty;
		}

		private static string NormalizeDocKey (string key)
		{
			string text = (key ?? string.Empty).Trim ();
			if (text.Length == 0) {
				return string.Empty;
			}
			long result;
			return long.TryParse (text, out result) ? result.ToString ("00") : text;
		}

		private static string JoinFields (params string[] fields)
		{
			for (int i = 0; i < fields.Length; i++) {
				fields [i] = EscapeField (fields [i] ?? string.Empty);
			}
			return string.Join ("\t", fields);
		}

		private static string EscapeField (string value)
		{
			return value.Replace ("\\", "\\\\").Replace ("\t", "\\t").Replace ("\r\n", "\\n")
				.Replace ("\r", "\\n")
				.Replace ("\n", "\\n");
		}

		private static string ColorToString (int red, int green, int blue)
		{
			return ColorTranslator.ToOle (Color.FromArgb (red, green, blue)).ToString ();
		}

		private static int CompareDocKeys (string leftKey, string rightKey)
		{
			if (long.TryParse (leftKey, out var result) && long.TryParse (rightKey, out var result2)) {
				return Math.Sign (result - result2);
			}
			return string.Compare (leftKey, rightKey, StringComparison.OrdinalIgnoreCase);
		}

		private static long GetTabBackColor (Worksheet masterSheet, string tabName, int lastRow)
		{
			string text = string.Empty;
			long result = 0L;
			for (int i = 3; i <= lastRow; i++) {
				string text2 = (Convert.ToString ((dynamic)(masterSheet.Cells [i, 5] as Range).Value2) ?? string.Empty).Trim ();
				if (text2.Length == 0) {
					text2 = "その他";
				}
				if (string.Equals (text2, tabName, StringComparison.OrdinalIgnoreCase)) {
					string text3 = KernelTemplateSyncService.NormalizeDocKey (Convert.ToString ((dynamic)(masterSheet.Cells [i, 1] as Range).Value2));
					if (text3.Length != 0 && (text.Length == 0 || CompareDocKeys (text3, text) < 0)) {
						text = text3;
						result = Convert.ToInt64 ((dynamic)(masterSheet.Cells [i, 6] as Range).Interior.Color);
					}
				}
			}
			return result;
		}

		private static int BuildPreferredPaneWidth (Worksheet masterSheet)
		{
			if (masterSheet == null) {
				return 720;
			}
			MasterSheetSnapshot masterSheetSnapshot = ReadMasterSheetSnapshot (masterSheet);
			int num = 0;
			int num2 = 0;
			for (int i = 0; i < masterSheetSnapshot.Rows.Count; i++) {
				string tabName = masterSheetSnapshot.Rows [i].TabName;
				string caption = masterSheetSnapshot.Rows [i].Caption;
				if (tabName.Length > num) {
					num = tabName.Length;
				}
				if (caption.Length > num2) {
					num2 = caption.Length;
				}
			}
			int num3 = 80 + num * 16 + num2 * 12;
			if (num3 < 420) {
				return 420;
			}
			if (num3 > 900) {
				return 900;
			}
			return num3;
		}

		private static MasterSheetSnapshot ReadMasterSheetSnapshot (Worksheet masterSheet)
		{
			if (masterSheet == null) {
				return new MasterSheetSnapshot (new List<MasterSheetRow> (), new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase));
			}
			Range range = null;
			try {
				int num = ((dynamic)masterSheet.Cells [masterSheet.Rows.Count, 1]).End [XlDirection.xlUp].Row;
				if (num < 3) {
					return new MasterSheetSnapshot (new List<MasterSheetRow> (), new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase));
				}
				range = masterSheet.Range [(dynamic)masterSheet.Cells [3, 1], (dynamic)masterSheet.Cells [num, 5]];
				if (!(range.Value2 is object[,] array)) {
					return new MasterSheetSnapshot (new List<MasterSheetRow> (), new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase));
				}
				List<MasterSheetRow> list = new List<MasterSheetRow> (array.GetUpperBound (0));
				Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
				Dictionary<string, long> dictionary2 = new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase);
				int upperBound = array.GetUpperBound (0);
				for (int i = 1; i <= upperBound; i++) {
					int rowIndex = 3 + i - 1;
					string text = NormalizeDocKey (Convert.ToString (array [i, 1]));
					string templateFileName = (Convert.ToString (array [i, 2]) ?? string.Empty).Trim ();
					string caption = (Convert.ToString (array [i, 3]) ?? string.Empty).Trim ();
					string text2 = (Convert.ToString (array [i, 5]) ?? string.Empty).Trim ();
					if (text2.Length == 0) {
						text2 = "その他";
					}
					long cellInteriorColor = GetCellInteriorColor (masterSheet, rowIndex, 4);
					long cellInteriorColor2 = GetCellInteriorColor (masterSheet, rowIndex, 6);
					list.Add (new MasterSheetRow (text, templateFileName, caption, text2, cellInteriorColor));
					if (text.Length != 0 && (!dictionary.TryGetValue (text2, out var value) || CompareDocKeys (text, value) < 0)) {
						dictionary [text2] = text;
						dictionary2 [text2] = cellInteriorColor2;
					}
				}
				return new MasterSheetSnapshot (list, dictionary2);
			} finally {
				ReleaseComObject (range);
			}
		}

		private static long GetCellInteriorColor (Worksheet worksheet, int rowIndex, int columnIndex)
		{
			Range range = null;
			Interior interior = null;
			try {
				range = worksheet.Cells [rowIndex, columnIndex] as Range;
				interior = range?.Interior;
				return Convert.ToInt64 ((dynamic)((interior == null) ? ((object)0) : interior.Color));
			} finally {
				ReleaseComObject (interior);
				ReleaseComObject (range);
			}
		}

		private static void ReleaseComObject (object comObject)
		{
			if (comObject == null) {
				return;
			}
			try {
				Marshal.FinalReleaseComObject (comObject);
			} catch {
			}
		}

		private AppState SaveAppState ()
		{
			return new AppState {
				ScreenUpdating = _application.ScreenUpdating,
				EnableEvents = _application.EnableEvents,
				DisplayAlerts = _application.DisplayAlerts
			};
		}

		private void ApplyFastSettings ()
		{
			_application.ScreenUpdating = false;
			_application.EnableEvents = false;
			_application.DisplayAlerts = false;
		}

		private void RestoreAppState (AppState state)
		{
			if (state == null) {
				return;
			}
			try {
				_application.ScreenUpdating = state.ScreenUpdating;
				_application.EnableEvents = state.EnableEvents;
				_application.DisplayAlerts = state.DisplayAlerts;
			} catch {
			}
		}

		private static SheetProtectionState SaveSheetProtectionState (Worksheet worksheet)
		{
			SheetProtectionState sheetProtectionState = new SheetProtectionState {
				IsProtected = (worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios)
			};
			if (sheetProtectionState.IsProtected) {
				Protection protection = worksheet.Protection;
				sheetProtectionState.AllowFormattingCells = protection.AllowFormattingCells;
				sheetProtectionState.AllowFormattingColumns = protection.AllowFormattingColumns;
				sheetProtectionState.AllowFormattingRows = protection.AllowFormattingRows;
				sheetProtectionState.AllowInsertingColumns = protection.AllowInsertingColumns;
				sheetProtectionState.AllowInsertingRows = protection.AllowInsertingRows;
				sheetProtectionState.AllowInsertingHyperlinks = protection.AllowInsertingHyperlinks;
				sheetProtectionState.AllowDeletingColumns = protection.AllowDeletingColumns;
				sheetProtectionState.AllowDeletingRows = protection.AllowDeletingRows;
				sheetProtectionState.AllowSorting = protection.AllowSorting;
				sheetProtectionState.AllowFiltering = protection.AllowFiltering;
				sheetProtectionState.AllowUsingPivotTables = protection.AllowUsingPivotTables;
			}
			return sheetProtectionState;
		}

		private static void RestoreSheetProtectionState (Worksheet worksheet, SheetProtectionState state)
		{
			if (worksheet == null || state == null || !state.IsProtected) {
				return;
			}
			try {
				worksheet.Protect (string.Empty, AllowFormattingCells: state.AllowFormattingCells, AllowFormattingColumns: state.AllowFormattingColumns, AllowFormattingRows: state.AllowFormattingRows, AllowInsertingColumns: state.AllowInsertingColumns, AllowInsertingRows: state.AllowInsertingRows, AllowInsertingHyperlinks: state.AllowInsertingHyperlinks, AllowDeletingColumns: state.AllowDeletingColumns, AllowDeletingRows: state.AllowDeletingRows, AllowSorting: state.AllowSorting, AllowFiltering: state.AllowFiltering, AllowUsingPivotTables: state.AllowUsingPivotTables, DrawingObjects: Type.Missing, Contents: Type.Missing, Scenarios: Type.Missing, UserInterfaceOnly: true);
				worksheet.EnableSelection = XlEnableSelection.xlUnlockedCells;
			} catch {
			}
		}
	}
}
