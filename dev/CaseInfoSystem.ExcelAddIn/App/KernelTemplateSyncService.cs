using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
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
		private sealed class PublicationExecutor
		{
			internal sealed class PublicationSideEffectResult
			{
				internal int UpdatedCount { get; }

				internal long MasterVersion { get; }

				internal bool BaseSyncSucceeded { get; }

				internal string BaseSyncError { get; }

				internal PublicationSideEffectResult (int updatedCount, long masterVersion, bool baseSyncSucceeded, string baseSyncError)
				{
					UpdatedCount = updatedCount;
					MasterVersion = masterVersion;
					BaseSyncSucceeded = baseSyncSucceeded;
					BaseSyncError = baseSyncError ?? string.Empty;
				}
			}

			private readonly Application _application;

			private readonly ExcelInteropService _excelInteropService;

			private readonly AccountingWorkbookService _accountingWorkbookService;

			private readonly PathCompatibilityService _pathCompatibilityService;

			private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

			private readonly MasterTemplateCatalogService _masterTemplateCatalogService;

			private readonly Logger _logger;

			internal PublicationExecutor (Application application, ExcelInteropService excelInteropService, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, MasterTemplateCatalogService masterTemplateCatalogService, Logger logger)
			{
				_application = application ?? throw new ArgumentNullException ("application");
				_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
				_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
				_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
				_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
				_masterTemplateCatalogService = masterTemplateCatalogService ?? throw new ArgumentNullException ("masterTemplateCatalogService");
				_logger = logger ?? throw new ArgumentNullException ("logger");
			}

			internal PublicationSideEffectResult PublishValidatedTemplates (Workbook kernelWorkbook, Worksheet masterSheet, string systemRoot, IReadOnlyList<TemplateRegistrationValidationEntry> templates)
			{
				int updatedCount = WriteToMasterList (masterSheet, templates);
				long masterVersion = IncrementTaskPaneMasterVersion (kernelWorkbook);
				// Kernel Save is the publication commit boundary.
				SaveKernelWorkbook (kernelWorkbook);
				string snapshotText = BuildTaskPaneSnapshot (masterSheet, systemRoot, masterVersion);
				string errorMessage;
				bool baseSyncSucceeded = TrySyncTaskPaneSnapshotToBase (systemRoot, snapshotText, masterVersion, out errorMessage);
				_masterTemplateCatalogService.InvalidateCache (kernelWorkbook);
				return new PublicationSideEffectResult (updatedCount, masterVersion, baseSyncSucceeded, errorMessage);
			}

			internal int WriteToMasterList (Worksheet masterSheet, IReadOnlyList<TemplateRegistrationValidationEntry> templates)
			{
				int num = MasterListFirstDataRow;
				int num2 = MasterListFirstDataRow + MasterListMaxKeyCount - 1;
				masterSheet.Range [(dynamic)masterSheet.Cells [num, ColumnA], (dynamic)masterSheet.Cells [num2, ColumnC]].ClearContents ();
				MasterListRowPayload masterListRowPayload = BuildMasterListRowPayload (templates);
				_accountingWorkbookService.WriteRangeValues (masterSheet, "$A$" + num.ToString () + ":$C$" + num2.ToString (), masterListRowPayload.Values);
				return masterListRowPayload.UpdatedCount;
			}

			private static MasterListRowPayload BuildMasterListRowPayload (IReadOnlyList<TemplateRegistrationValidationEntry> templates)
			{
				object[,] array = new object[MasterListMaxKeyCount, 3];
				int num = 0;
				if (templates == null) {
					return new MasterListRowPayload (array, num);
				}
				foreach (TemplateRegistrationValidationEntry template in templates) {
					if (!TryBuildMasterListRow (template, array)) {
						continue;
					}
					num++;
				}
				return new MasterListRowPayload (array, num);
			}

			private static bool TryBuildMasterListRow (TemplateRegistrationValidationEntry template, object[,] values)
			{
				if (template == null || !int.TryParse (template.Key, out var result) || result < 1 || result > MasterListMaxKeyCount) {
					return false;
				}
				string text = template.FileName ?? string.Empty;
				int num = result - 1;
				values [num, ColumnA - 1] = result.ToString ("00");
				values [num, ColumnB - 1] = text;
				values [num, ColumnC - 1] = template.DisplayName ?? KernelTemplateSyncService.ExtractDocumentName (text);
				return true;
			}

			private sealed class MasterListRowPayload
			{
				internal object[,] Values { get; }

				internal int UpdatedCount { get; }

				internal MasterListRowPayload (object[,] values, int updatedCount)
				{
					Values = values ?? throw new ArgumentNullException ("values");
					UpdatedCount = updatedCount;
				}
			}

			internal long IncrementTaskPaneMasterVersion (Workbook kernelWorkbook)
			{
				string existingVersion = _excelInteropService.TryGetDocumentProperty (kernelWorkbook, TaskPaneMasterVersionProp);
				long nextVersion = KernelTemplateSyncService.CreateNextTaskPaneMasterVersion (existingVersion, DateTime.Today);
				_excelInteropService.SetDocumentProperty (kernelWorkbook, TaskPaneMasterVersionProp, nextVersion.ToString (CultureInfo.InvariantCulture));
				return nextVersion;
			}

			internal void SaveKernelWorkbook (Workbook kernelWorkbook)
			{
				using (var saveScope = new ExcelApplicationStateScope (_application, suppressRestoreExceptions: true)) {
					saveScope.SetDisplayAlerts (false);
					kernelWorkbook.Save ();
				}
			}

			internal bool TrySyncTaskPaneSnapshotToBase (string systemRoot, string snapshotText, long masterVersion, out string errorMessage)
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
								WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave (
									workbook,
									_logger,
									"KernelTemplateSyncService.CloseOpenedBaseWorkbook");
							}
						}
					}
				}
			}

			internal string BuildTaskPaneSnapshot (Worksheet masterSheet, string systemRoot, long masterVersion)
			{
				List<string> list = new List<string> ();
				string text = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath (systemRoot, _pathCompatibilityService);
				string text2 = Path.GetFileName (text);
				if (string.IsNullOrWhiteSpace (text2)) {
					text2 = WorkbookFileNameResolver.BuildBaseWorkbookName (Path.GetExtension (text));
				}
				list.Add (KernelTemplateSyncService.JoinFields ("META", "2", text2, text, KernelTemplateSyncService.BuildPreferredPaneWidth (masterSheet).ToString (), masterVersion.ToString (CultureInfo.InvariantCulture)));
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

			private void SaveSnapshotToBaseWorkbook (Workbook baseWorkbook, string snapshotText, long masterVersion)
			{
				BaseSnapshotPropertyStorage.Save (_excelInteropService, baseWorkbook, snapshotText, masterVersion);
			}

			private static class BaseSnapshotPropertyStorage
			{
				internal static void Save (ExcelInteropService excelInteropService, Workbook baseWorkbook, string snapshotText, long masterVersion)
				{
					string s = excelInteropService.TryGetDocumentProperty (baseWorkbook, TaskPaneBaseCacheCountProp);
					int result;
					int num = (int.TryParse (s, out result) ? result : 0);
					if (string.IsNullOrEmpty (snapshotText)) {
						excelInteropService.SetDocumentProperty (baseWorkbook, TaskPaneBaseCacheCountProp, "0");
						excelInteropService.SetDocumentProperty (baseWorkbook, TaskPaneBaseMasterVersionProp, masterVersion.ToString (CultureInfo.InvariantCulture));
						return;
					}
					int num2 = CalculateChunkCount (snapshotText);
					excelInteropService.SetDocumentProperty (baseWorkbook, TaskPaneBaseCacheCountProp, num2.ToString ());
					excelInteropService.SetDocumentProperty (baseWorkbook, TaskPaneBaseMasterVersionProp, masterVersion.ToString (CultureInfo.InvariantCulture));
					excelInteropService.SetDocumentProperty (baseWorkbook, TaskPaneMasterVersionProp, masterVersion.ToString (CultureInfo.InvariantCulture));
					WriteSnapshotChunks (excelInteropService, baseWorkbook, snapshotText, num2);
					ClearStaleSnapshotChunks (excelInteropService, baseWorkbook, num2 + 1, num);
				}

				private static int CalculateChunkCount (string snapshotText)
				{
					return (snapshotText.Length - 1) / TaskPaneCacheChunkSize + 1;
				}

				private static void WriteSnapshotChunks (ExcelInteropService excelInteropService, Workbook baseWorkbook, string snapshotText, int chunkCount)
				{
					for (int i = 1; i <= chunkCount; i++) {
						int num = (i - 1) * TaskPaneCacheChunkSize;
						int length = Math.Min (TaskPaneCacheChunkSize, snapshotText.Length - num);
						excelInteropService.SetDocumentProperty (baseWorkbook, BuildChunkPropertyName (i), snapshotText.Substring (num, length));
					}
				}

				private static void ClearStaleSnapshotChunks (ExcelInteropService excelInteropService, Workbook baseWorkbook, int firstStaleChunkIndex, int previousChunkCount)
				{
					for (int i = firstStaleChunkIndex; i <= previousChunkCount; i++) {
						excelInteropService.SetDocumentProperty (baseWorkbook, BuildChunkPropertyName (i), string.Empty);
					}
				}

				private static string BuildChunkPropertyName (int chunkIndex)
				{
					return TaskPaneBaseCachePartPropPrefix + chunkIndex.ToString ("00");
				}
			}
		}

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

		private const long TaskPaneVersionSequenceBase = 1000L;

		private const long TaskPaneMaxDailySequence = 999L;

		private const string AllTabCaption = "全て";

		private const string DefaultTabCaption = "その他";

		private const string CaseListActionCaption = "案件一覧登録（未了）";

		private const string AccountingActionCaption = "会計書類セット";

		private readonly Application _application;

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

		private readonly KernelTemplateSyncPreparationService _preparationService;

		private readonly PublicationExecutor _publicationExecutor;

		private readonly Logger _logger;

		internal KernelTemplateSyncService (Application application, KernelWorkbookService kernelWorkbookService, ExcelInteropService excelInteropService, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, KernelTemplateSyncPreparationService preparationService, MasterTemplateCatalogService masterTemplateCatalogService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			if (excelInteropService == null) {
				throw new ArgumentNullException ("excelInteropService");
			}
			if (accountingWorkbookService == null) {
				throw new ArgumentNullException ("accountingWorkbookService");
			}
			if (pathCompatibilityService == null) {
				throw new ArgumentNullException ("pathCompatibilityService");
			}
			_preparationService = preparationService ?? throw new ArgumentNullException ("preparationService");
			_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_publicationExecutor = new PublicationExecutor (application, excelInteropService, accountingWorkbookService, pathCompatibilityService, caseWorkbookLifecycleService, masterTemplateCatalogService, logger);
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
				using (KernelTemplateSyncPreparationService.PreparedKernelTemplateSyncScope preparedScope = _preparationService.Prepare (openKernelWorkbook)) {
					KernelTemplateSyncPreflightResult kernelTemplateSyncPreflightResult = preparedScope.PreflightResult;
					if (kernelTemplateSyncPreflightResult.Status != KernelTemplateSyncPreflightStatus.Succeeded) {
						return CreatePreflightFailureResult (kernelTemplateSyncPreflightResult);
					}
					string text = kernelTemplateSyncPreflightResult.TemplateDirectory;
					TemplateRegistrationValidationSummary templateRegistrationValidationSummary = kernelTemplateSyncPreflightResult.ValidationSummary;
					IReadOnlyList<TemplateRegistrationValidationEntry> validTemplates = templateRegistrationValidationSummary.GetValidTemplates ();
					PublicationExecutor.PublicationSideEffectResult publicationResult = _publicationExecutor.PublishValidatedTemplates (openKernelWorkbook, preparedScope.MasterSheet, preparedScope.SystemRoot, validTemplates);
					_logger.Info ("Kernel template sync completed. updatedCount=" + publicationResult.UpdatedCount + ", detectedCount=" + templateRegistrationValidationSummary.DetectedFileCount + ", excludedCount=" + templateRegistrationValidationSummary.ExcludedTemplateCount + ", warningCount=" + templateRegistrationValidationSummary.WarningFileCount + ", masterVersion=" + publicationResult.MasterVersion);
					return new KernelTemplateSyncResult {
						Success = true,
						UpdatedCount = publicationResult.UpdatedCount,
						DetectedCount = templateRegistrationValidationSummary.DetectedFileCount,
						ExcludedCount = templateRegistrationValidationSummary.ExcludedTemplateCount,
						WarningCount = templateRegistrationValidationSummary.WarningFileCount,
						MasterVersion = publicationResult.MasterVersion,
						TemplateDirectory = text,
						TemplateResults = templateRegistrationValidationSummary.TemplateResults,
						BaseSyncError = publicationResult.BaseSyncError,
						Message = BuildCompletedMessage (text, publicationResult.UpdatedCount, templateRegistrationValidationSummary, stopwatch.Elapsed, publicationResult.MasterVersion, publicationResult.BaseSyncSucceeded, publicationResult.BaseSyncError)
					};
				}
			}
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

		internal static long CreateNextTaskPaneMasterVersion (string existingVersionText, DateTime today)
		{
			long existingVersion = ParsePositiveLong (existingVersionText);
			long todayKey = BuildTaskPaneVersionDateKey (today);
			long candidateVersion = todayKey * TaskPaneVersionSequenceBase + 1L;
			long existingDateKey;
			long existingSequence;
			if (TryGetTaskPaneVersionParts (existingVersion, out existingDateKey, out existingSequence) && existingDateKey == todayKey) {
				if (existingSequence >= TaskPaneMaxDailySequence) {
					throw new InvalidOperationException ("TaskPane master version daily sequence exceeded 999.");
				}
				candidateVersion = todayKey * TaskPaneVersionSequenceBase + existingSequence + 1L;
			}
			if (existingVersion > 0L && candidateVersion <= existingVersion) {
				if (existingVersion == long.MaxValue) {
					throw new InvalidOperationException ("TaskPane master version cannot be incremented safely.");
				}
				candidateVersion = existingVersion + 1L;
			}
			return candidateVersion;
		}

		private static long ParsePositiveLong (string value)
		{
			long result;
			return long.TryParse (value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result) && result > 0L ? result : 0L;
		}

		private static long BuildTaskPaneVersionDateKey (DateTime today)
		{
			DateTime date = today.Date;
			return date.Year * 10000L + date.Month * 100L + date.Day;
		}

		private static bool TryGetTaskPaneVersionParts (long version, out long dateKey, out long sequence)
		{
			dateKey = 0L;
			sequence = 0L;
			if (version <= 0L) {
				return false;
			}
			dateKey = version / TaskPaneVersionSequenceBase;
			sequence = version % TaskPaneVersionSequenceBase;
			return dateKey >= 10000101L && dateKey <= 99991231L && sequence >= 1L && sequence <= TaskPaneMaxDailySequence;
		}

		private static string BuildCompletedMessage (string templateDirectory, int updatedCount, TemplateRegistrationValidationSummary validationSummary, TimeSpan elapsed, long masterVersion, bool baseSyncSucceeded, string baseSyncError)
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
				.Append (masterVersion.ToString (CultureInfo.InvariantCulture));
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
	}
}
