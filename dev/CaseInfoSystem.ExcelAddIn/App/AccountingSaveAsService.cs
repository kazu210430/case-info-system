using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingSaveAsService
	{
		private sealed class ExcelApplicationState
		{
			private readonly Microsoft.Office.Interop.Excel.Application _application;

			private bool ScreenUpdating { get; }

			private bool EnableEvents { get; }

			private bool DisplayAlerts { get; }

			private ExcelApplicationState (Microsoft.Office.Interop.Excel.Application application, bool screenUpdating, bool enableEvents, bool displayAlerts)
			{
				_application = application;
				ScreenUpdating = screenUpdating;
				EnableEvents = enableEvents;
				DisplayAlerts = displayAlerts;
			}

			internal static ExcelApplicationState CaptureAndApply (Microsoft.Office.Interop.Excel.Application application)
			{
				if (application == null) {
					return null;
				}
				ExcelApplicationState result = new ExcelApplicationState (application, application.ScreenUpdating, application.EnableEvents, application.DisplayAlerts);
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				application.DisplayAlerts = false;
				return result;
			}

			internal void Restore ()
			{
				if (_application != null) {
					_application.ScreenUpdating = ScreenUpdating;
					_application.EnableEvents = EnableEvents;
					_application.DisplayAlerts = DisplayAlerts;
				}
			}
		}

		private const string ProcedureName = "AccountingSaveAs";

		private const string RoleDocumentPropertyName = "ROLE";

		private const string CaseRoleValue = "CASE";

		private const string DefaultNameRuleA = "YY";

		private const string DefaultNameRuleB = "DOC_CUST";

		private const string CaseHomeSheetCodeName = "shHOME";

		private const string CaseHomeSheetName = "ホーム";

		private const string CustomerNameKey = "顧客_名前";

		private readonly ExcelInteropService _excelInteropService;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly DocumentOutputService _documentOutputService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly UserErrorService _userErrorService;

		private readonly Logger _logger;

		internal AccountingSaveAsService (ExcelInteropService excelInteropService, AccountingWorkbookService accountingWorkbookService, DocumentOutputService documentOutputService, PathCompatibilityService pathCompatibilityService, UserErrorService userErrorService, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_documentOutputService = documentOutputService ?? throw new ArgumentNullException ("documentOutputService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_userErrorService = userErrorService ?? throw new ArgumentNullException ("userErrorService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (WorkbookContext context)
		{
			if (context == null) {
				throw new ArgumentNullException ("context");
			}
			if (context.Workbook == null) {
				throw new InvalidOperationException ("保存対象のWorkbookを取得できませんでした。");
			}
			Workbook workbook = null;
			bool openedTemporarily = false;
			ExcelApplicationState excelApplicationState = null;
			try {
				excelApplicationState = ExcelApplicationState.CaptureAndApply (context.Workbook.Application);
				workbook = ResolveCaseWorkbook (context.Workbook, out openedTemporarily);
				if (workbook == null) {
					throw new InvalidOperationException ("CASEブックを取得できませんでした。");
				}
				string documentName = ResolveActiveSheetName (context.Workbook);
				string customerName = ResolveCustomerName (workbook);
				string right = BuildOutputFileName (workbook, documentName, customerName);
				string text = _documentOutputService.ResolveWorkbookFolder (workbook);
				if (string.IsNullOrWhiteSpace (text)) {
					throw new InvalidOperationException ("保存先フォルダのローカルパスを解決できませんでした。");
				}
				string rawFullPath = _pathCompatibilityService.CombinePath (text, right);
				string text2 = _documentOutputService.PrepareSavePath (rawFullPath);
				if (string.IsNullOrWhiteSpace (text2)) {
					throw new InvalidOperationException ("保存先ファイルパスを確定できませんでした。");
				}
				_logger.Info ("Accounting save-as started. workbook=" + (_excelInteropService.GetWorkbookFullName (context.Workbook) ?? string.Empty) + ", caseWorkbook=" + (_excelInteropService.GetWorkbookFullName (workbook) ?? string.Empty) + ", savePath=" + text2);
				_accountingWorkbookService.SaveAsMacroEnabled (context.Workbook, text2);
				MessageBox.Show ("保存しました。" + Environment.NewLine + text2, "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingSaveAs", exception);
			} finally {
				if (openedTemporarily && workbook != null) {
					try {
						_accountingWorkbookService.CloseWithoutSaving (workbook);
					} catch (Exception exception2) {
						_logger.Error ("Accounting save-as temporary case workbook close failed.", exception2);
					}
				}
				excelApplicationState?.Restore ();
			}
		}

		private Workbook ResolveCaseWorkbook (Workbook accountingWorkbook, out bool openedTemporarily)
		{
			openedTemporarily = false;
			string sourceCasePath = _excelInteropService.TryGetDocumentProperty (accountingWorkbook, "SOURCE_CASE_PATH");
			Workbook workbook = ResolveCaseWorkbookByPath (sourceCasePath, out openedTemporarily);
			if (workbook != null) {
				return workbook;
			}
			return FindCaseWorkbookInSameFolder (accountingWorkbook, out openedTemporarily);
		}

		private string ResolveCustomerName (Workbook caseWorkbook)
		{
			if (caseWorkbook == null) {
				return string.Empty;
			}
			Worksheet worksheet = null;
			try {
				worksheet = _excelInteropService.FindWorksheetByCodeName (caseWorkbook, "shHOME");
				if (worksheet == null) {
					worksheet = caseWorkbook.Worksheets ["ホーム"] as Worksheet;
				}
				IReadOnlyDictionary<string, string> readOnlyDictionary = _excelInteropService.ReadKeyValueMapFromColumnsAandB (worksheet);
				if (readOnlyDictionary == null) {
					return string.Empty;
				}
				if (!readOnlyDictionary.TryGetValue ("顧客_名前", out var value)) {
					return string.Empty;
				}
				return (value ?? string.Empty).Trim ();
			} finally {
				ReleaseComObject (worksheet);
			}
		}

		private Workbook ResolveCaseWorkbookByPath (string sourceCasePath, out bool openedTemporarily)
		{
			openedTemporarily = false;
			string text = _pathCompatibilityService.ResolveToExistingLocalPath (sourceCasePath);
			if (string.IsNullOrWhiteSpace (text)) {
				return null;
			}
			Workbook workbook = _excelInteropService.FindOpenWorkbook (text);
			if (IsCaseWorkbook (workbook)) {
				return workbook;
			}
			if (!_pathCompatibilityService.FileExistsSafe (text)) {
				return null;
			}
			Workbook workbook2 = _accountingWorkbookService.OpenReadOnlyHiddenInCurrentApplication (text);
			if (!IsCaseWorkbook (workbook2)) {
				if (workbook2 != null) {
					_accountingWorkbookService.CloseWithoutSaving (workbook2);
				}
				return null;
			}
			openedTemporarily = true;
			return workbook2;
		}

		private Workbook FindCaseWorkbookInSameFolder (Workbook accountingWorkbook, out bool openedTemporarily)
		{
			openedTemporarily = false;
			string text = _pathCompatibilityService.ResolveToExistingLocalPath (_excelInteropService.GetWorkbookPath (accountingWorkbook));
			if (string.IsNullOrWhiteSpace (text)) {
				return null;
			}
			string b = _pathCompatibilityService.ResolveToExistingLocalPath (_excelInteropService.GetWorkbookFullName (accountingWorkbook));
			string[] caseWorkbookCandidatePaths;
			try {
				caseWorkbookCandidatePaths = GetCaseWorkbookCandidatePaths (text);
			} catch (Exception innerException) {
				throw new InvalidOperationException ("CASE探索用フォルダのファイル一覧取得に失敗しました。", innerException);
			}
			for (int i = 0; i < caseWorkbookCandidatePaths.Length; i++) {
				string text2 = _pathCompatibilityService.ResolveToExistingLocalPath (caseWorkbookCandidatePaths [i]);
				if (!string.IsNullOrWhiteSpace (text2) && !string.Equals (text2, b, StringComparison.OrdinalIgnoreCase)) {
					bool openedTemporarily2 = false;
					Workbook workbook = ResolveCaseWorkbookByPath (text2, out openedTemporarily2);
					if (workbook != null) {
						openedTemporarily = openedTemporarily2;
						return workbook;
					}
				}
			}
			return null;
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

		private string[] GetCaseWorkbookCandidatePaths (string accountingFolderPath)
		{
			List<string> list = new List<string> ();
			string[] supportedMainWorkbookExtensions = WorkbookFileNameResolver.GetSupportedMainWorkbookExtensions ();
			for (int i = 0; i < supportedMainWorkbookExtensions.Length; i++) {
				string[] files = Directory.GetFiles (accountingFolderPath, "*" + supportedMainWorkbookExtensions [i]);
				if (files.Length != 0) {
					list.AddRange (files);
				}
			}
			return list.ToArray ();
		}

		private string BuildOutputFileName (Workbook caseWorkbook, string documentName, string customerName)
		{
			string text = _excelInteropService.TryGetDocumentProperty (caseWorkbook, "NAME_RULE_A");
			string text2 = _excelInteropService.TryGetDocumentProperty (caseWorkbook, "NAME_RULE_B");
			if (string.IsNullOrWhiteSpace (text)) {
				text = "YY";
			}
			if (string.IsNullOrWhiteSpace (text2)) {
				text2 = "DOC_CUST";
			}
			string text3 = KernelNamingService.BuildDocumentName (text, text2, (documentName ?? string.Empty).Trim (), (customerName ?? string.Empty).Trim (), DateTime.Today);
			if (string.IsNullOrWhiteSpace (text3)) {
				throw new InvalidOperationException ("保存ファイル名を生成できませんでした。");
			}
			return text3.Trim () + _documentOutputService.ResolveWorkbookOutputExtension (caseWorkbook);
		}

		private static string ResolveActiveSheetName (Workbook workbook)
		{
			Worksheet worksheet = ((workbook == null) ? null : (workbook.ActiveSheet as Worksheet));
			try {
				string text = ((worksheet == null) ? string.Empty : (worksheet.Name ?? string.Empty));
				if (string.IsNullOrWhiteSpace (text)) {
					throw new InvalidOperationException ("現在シート名を取得できませんでした。");
				}
				return text.Trim ();
			} finally {
				if (worksheet != null) {
					try {
						Marshal.ReleaseComObject (worksheet);
					} catch {
					}
				}
			}
		}

		private bool IsCaseWorkbook (Workbook workbook)
		{
			if (workbook == null) {
				return false;
			}
			string text = _excelInteropService.TryGetDocumentProperty (workbook, "ROLE");
			return string.Equals ((text ?? string.Empty).Trim (), "CASE", StringComparison.OrdinalIgnoreCase);
		}
	}
}
