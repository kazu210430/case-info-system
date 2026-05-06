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

		private sealed class PublicationExecutor
		{
			private readonly Application _application;

			private readonly ExcelInteropService _excelInteropService;

			private readonly AccountingWorkbookService _accountingWorkbookService;

			private readonly PathCompatibilityService _pathCompatibilityService;

			private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

			private readonly Logger _logger;

			internal PublicationExecutor (Application application, ExcelInteropService excelInteropService, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, Logger logger)
			{
				_application = application ?? throw new ArgumentNullException ("application");
				_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
				_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
				_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
				_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
				_logger = logger ?? throw new ArgumentNullException ("logger");
			}

			internal int WriteToMasterList (Worksheet masterSheet, IReadOnlyList<TemplateRegistrationValidationEntry> templates)
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
						array [num5, 2] = template.DisplayName ?? KernelTemplateSyncService.ExtractDocumentName (text);
						num3++;
					}
				}
				_accountingWorkbookService.WriteRangeValues (masterSheet, "$A$" + num.ToString () + ":$C$" + num2.ToString (), array);
				return num3;
			}

			internal int IncrementTaskPaneMasterVersion (Workbook kernelWorkbook)
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

			internal void SaveKernelWorkbook (Workbook kernelWorkbook)
			{
				using (var saveScope = new ExcelApplicationStateScope (_application, suppressRestoreExceptions: true)) {
					saveScope.SetDisplayAlerts (false);
					kernelWorkbook.Save ();
				}
			}

			internal bool TrySyncTaskPaneSnapshotToBase (string systemRoot, string snapshotText, int masterVersion, out string errorMessage)
			{
				errorMessage = string.Empty;
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
					using (var saveScope = new ExcelApplicationStateScope (_application, suppressRestoreExceptions: true)) {
						saveScope.SetDisplayAlerts (false);
						workbook.Save ();
					}
					return true;
				} catch (Exception ex) {
					errorMessage = ex.Message;
					_logger.Error ("Kernel template sync base snapshot failed.", ex);
					return false;
				} finally {
					if (!flag && workbook != null) {
						using (var closeScope = new ExcelApplicationStateScope (_application, suppressRestoreExceptions: true)) {
							closeScope.SetDisplayAlerts (false);
							using (_caseWorkbookLifecycleService.BeginManagedCloseScope (workbook)) {
								workbook.Close (false, Type.Missing, Type.Missing);
							}
						}
					}
				}
			}

			internal string BuildTaskPaneSnapshot (Worksheet masterSheet, string systemRoot, int masterVersion)
			{
				List<string> list = new List<string> ();
				string text = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath (systemRoot, _pathCompatibilityService);
				string text2 = Path.GetFileName (text);
				if (string.IsNullOrWhiteSpace (text2)) {
					text2 = WorkbookFileNameResolver.BuildBaseWorkbookName (Path.GetExtension (text));
				}
				list.Add (KernelTemplateSyncService.JoinFields ("META", "2", text2, text, KernelTemplateSyncService.BuildPreferredPaneWidth (masterSheet).ToString (), masterVersion.ToString ()));
				list.Add (KernelTemplateSyncService.JoinFields ("SPECIAL", "btnCaseList", "案件一覧登録（未了）", "caselist", string.Empty, "18", "16", "128", "32", KernelTemplateSyncService.ColorToString (248, 225, 193)));
				list.Add (KernelTemplateSyncService.JoinFields ("SPECIAL", "btnAccounting", "会計書類セット", "accounting", string.Empty, "18", "64", "128", "32", KernelTemplateSyncService.ColorToString (226, 239, 218)));
				Dictionary<string, int> dictionary = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
				MasterTemplateSheetData masterSheetSnapshot = MasterTemplateSheetReader.Read (masterSheet);
				Dictionary<string, long> tabBackColors = KernelTemplateSyncService.BuildTabBackColors (masterSheetSnapshot.Rows);
				Dictionary<string, int> dictionary2 = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
				for (int i = 0; i < masterSheetSnapshot.Rows.Count; i++) {
					MasterTemplateSheetRowData masterSheetRow = masterSheetSnapshot.Rows [i];
					string key = masterSheetRow.Key;
					string templateFileName = masterSheetRow.TemplateFileName;
					string caption = masterSheetRow.Caption;
					string tabName = KernelTemplateSyncService.NormalizeTabName (masterSheetRow.TabName);
					if (key.Length != 0 && caption.Length != 0) {
						if (!dictionary.ContainsKey (tabName)) {
							int value = dictionary.Count + 1;
							dictionary.Add (tabName, value);
							dictionary2 [tabName] = 0;
							long tabBackColor = KernelTemplateSyncService.GetTabBackColor (tabBackColors, tabName);
							list.Add (KernelTemplateSyncService.JoinFields ("TAB", value.ToString (), tabName, tabBackColor.ToString ()));
						}
						int num = (dictionary2 [tabName] = ((!dictionary2.ContainsKey (tabName)) ? 1 : (dictionary2 [tabName] + 1)));
						list.Add (KernelTemplateSyncService.JoinFields ("DOC", "btnDoc_" + key, key, caption, "doc", tabName, num.ToString (), masterSheetRow.FillColor.ToString (), templateFileName));
					}
				}
				if (!dictionary.ContainsKey ("全て")) {
					int value2 = dictionary.Count + 1;
					dictionary.Add ("全て", value2);
					list.Add (KernelTemplateSyncService.JoinFields ("TAB", value2.ToString (), "全て", ColorTranslator.ToOle (Color.FromArgb (255, 255, 255)).ToString ()));
				}
				return string.Join (Environment.NewLine, list.ToArray ());
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

		private readonly KernelTemplateSyncPreflightService _kernelTemplateSyncPreflightService;

		private readonly MasterTemplateCatalogService _masterTemplateCatalogService;

		private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

		private readonly PublicationExecutor _publicationExecutor;

		private readonly Logger _logger;

		internal KernelTemplateSyncService (Application application, KernelWorkbookService kernelWorkbookService, ExcelInteropService excelInteropService, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository, KernelTemplateSyncPreflightService kernelTemplateSyncPreflightService, MasterTemplateCatalogService masterTemplateCatalogService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException ("caseListFieldDefinitionRepository");
			_kernelTemplateSyncPreflightService = kernelTemplateSyncPreflightService ?? throw new ArgumentNullException ("kernelTemplateSyncPreflightService");
			_masterTemplateCatalogService = masterTemplateCatalogService ?? throw new ArgumentNullException ("masterTemplateCatalogService");
			_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_publicationExecutor = new PublicationExecutor (application, excelInteropService, accountingWorkbookService, pathCompatibilityService, caseWorkbookLifecycleService, logger);
		}

		internal KernelTemplateSyncResult Execute (WorkbookContext context)
		{
			if (context == null) {
				throw new InvalidOperationException ("WorkbookContext is required for template sync.");
			}
			Workbook openKernelWorkbook = _kernelWorkbookService.ResolveKernelWorkbook (context);
			if (openKernelWorkbook == null) {
				throw new InvalidOperationException ("Kernel ブックを開いてから実行してください。");
			}
			Stopwatch stopwatch = Stopwatch.StartNew ();
			using (var excelApplicationStateScope = new ExcelApplicationStateScope (_application, suppressRestoreExceptions: true)) {
				excelApplicationStateScope.SetScreenUpdating (false);
				excelApplicationStateScope.SetEnableEvents (false);
				Worksheet worksheet = null;
				SheetProtectionState sheetProtectionState = null;
				try {
					worksheet = GetMasterListSheet (openKernelWorkbook);
					sheetProtectionState = SaveSheetProtectionState (worksheet);
					if (sheetProtectionState.IsProtected) {
						worksheet.Unprotect (string.Empty);
					}
					ValidateMasterListSheet (worksheet);
					string systemRoot = ResolveSystemRoot (openKernelWorkbook);
					KernelTemplateSyncPreflightResult kernelTemplateSyncPreflightResult = _kernelTemplateSyncPreflightService.Run (new KernelTemplateSyncPreflightRequest (systemRoot, LoadDefinedTemplateTags (openKernelWorkbook)));
					if (kernelTemplateSyncPreflightResult.Status != KernelTemplateSyncPreflightStatus.Succeeded) {
						return CreatePreflightFailureResult (kernelTemplateSyncPreflightResult);
					}
					string text = kernelTemplateSyncPreflightResult.TemplateDirectory;
					TemplateRegistrationValidationSummary templateRegistrationValidationSummary = kernelTemplateSyncPreflightResult.ValidationSummary;
					IReadOnlyList<TemplateRegistrationValidationEntry> validTemplates = templateRegistrationValidationSummary.GetValidTemplates ();
					int updatedCount = _publicationExecutor.WriteToMasterList (worksheet, validTemplates);
					int masterVersion = _publicationExecutor.IncrementTaskPaneMasterVersion (openKernelWorkbook);
					// Kernel Save is the publication commit boundary.
					_publicationExecutor.SaveKernelWorkbook (openKernelWorkbook);
					string snapshotText = _publicationExecutor.BuildTaskPaneSnapshot (worksheet, systemRoot, masterVersion);
					string errorMessage;
					bool baseSyncSucceeded = _publicationExecutor.TrySyncTaskPaneSnapshotToBase (systemRoot, snapshotText, masterVersion, out errorMessage);
					_masterTemplateCatalogService.InvalidateCache (openKernelWorkbook);
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
				}
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

		private static KernelTemplateSyncResult CreatePreflightFailureResult (KernelTemplateSyncPreflightResult preflightResult)
		{
			ValidationFailureSummary failure = preflightResult?.Failure;
			return new KernelTemplateSyncResult {
				Success = false,
				TemplateDirectory = preflightResult?.TemplateDirectory ?? string.Empty,
				DetectedCount = failure?.DetectedCount ?? 0,
				TemplateResults = failure?.TemplateResults ?? Array.Empty<TemplateRegistrationValidationEntry> (),
				Message = failure?.Message ?? string.Empty
			};
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

		private static string NormalizeTabName (string tabName)
		{
			string text = (tabName ?? string.Empty).Trim ();
			return (text.Length != 0) ? text : "その他";
		}

		private static Dictionary<string, long> BuildTabBackColors (IReadOnlyList<MasterTemplateSheetRowData> rows)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			Dictionary<string, long> dictionary2 = new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase);
			if (rows == null) {
				return dictionary2;
			}
			for (int i = 0; i < rows.Count; i++) {
				MasterTemplateSheetRowData masterTemplateSheetRowData = rows [i];
				string key = masterTemplateSheetRowData.Key;
				string text = NormalizeTabName (masterTemplateSheetRowData.TabName);
				if (key.Length != 0 && (!dictionary.TryGetValue (text, out var value) || CompareDocKeys (key, value) < 0)) {
					dictionary [text] = key;
					dictionary2 [text] = masterTemplateSheetRowData.TabBackColor;
				}
			}
			return dictionary2;
		}

		private static long GetTabBackColor (IReadOnlyDictionary<string, long> tabBackColors, string tabName)
		{
			if (tabBackColors == null || string.IsNullOrWhiteSpace (tabName)) {
				return 0L;
			}
			long value;
			return tabBackColors.TryGetValue (tabName, out value) ? value : 0L;
		}

		private static int BuildPreferredPaneWidth (Worksheet masterSheet)
		{
			if (masterSheet == null) {
				return 720;
			}
			MasterTemplateSheetData masterSheetSnapshot = MasterTemplateSheetReader.Read (masterSheet);
			int num = 0;
			int num2 = 0;
			for (int i = 0; i < masterSheetSnapshot.Rows.Count; i++) {
				string tabName = NormalizeTabName (masterSheetSnapshot.Rows [i].TabName);
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
