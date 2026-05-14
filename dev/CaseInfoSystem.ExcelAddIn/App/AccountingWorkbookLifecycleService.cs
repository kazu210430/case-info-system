using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingWorkbookLifecycleService
	{
		private const string CurrencyRangeAddress = "F15:F33";

		private const string CurrencyNumberFormatLocal = "\\#,##0;\\-#,##0";

		private const string TopLeftCellAddress = "A1";

		private const string ArgumentSheetName = "引数";

		private readonly WorkbookRoleResolver _workbookRoleResolver;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly AccountingFormHelperService _accountingFormHelperService;

		private readonly AccountingPaymentHistoryImportService _accountingPaymentHistoryImportService;

		private readonly Logger _logger;

		private readonly HashSet<string> _initializedWorkbookKeys;

		private readonly HashSet<string> _activeSheetSynchronizedWorkbookKeys;

		internal AccountingWorkbookLifecycleService (WorkbookRoleResolver workbookRoleResolver, AccountingWorkbookService accountingWorkbookService, AccountingFormHelperService accountingFormHelperService, AccountingPaymentHistoryImportService accountingPaymentHistoryImportService, Logger logger)
		{
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_accountingFormHelperService = accountingFormHelperService ?? throw new ArgumentNullException ("accountingFormHelperService");
			_accountingPaymentHistoryImportService = accountingPaymentHistoryImportService ?? throw new ArgumentNullException ("accountingPaymentHistoryImportService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_initializedWorkbookKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
			_activeSheetSynchronizedWorkbookKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
		}

		internal void HandleWorkbookOpenedOrActivated (Workbook workbook, string eventName)
		{
			bool isAccountingWorkbook = _workbookRoleResolver.IsAccountingWorkbook (workbook);
			if (!isAccountingWorkbook) {
				return;
			}
			EnsureWorkbookInitialized (workbook);
			if (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet (eventName, isAccountingWorkbook)) {
				SynchronizeActiveSheetFromWindowActivation (workbook, ResolveActiveSheetWindow (workbook), eventName);
			}
		}

		internal void HandleWindowActivated (Workbook workbook, Window window, string eventName)
		{
			bool isAccountingWorkbook = _workbookRoleResolver.IsAccountingWorkbook (workbook);
			if (!AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet (eventName, isAccountingWorkbook)) {
				return;
			}
			if (window == null) {
				_logger.Info ("Accounting workbook active sheet sync skipped. reason=" + (eventName ?? string.Empty) + ", eventWindowMissing=True, workbook=" + GetWorkbookKey (workbook));
				return;
			}
			EnsureWorkbookInitialized (workbook);
			SynchronizeActiveSheetFromWindowActivation (workbook, window, eventName);
		}

		internal void HandleSheetActivated (object sheetObject)
		{
			if (!(sheetObject is Worksheet worksheet)) {
				return;
			}
			HandleActivatedWorksheet (worksheet, "SheetActivate", null);
		}

		private void SynchronizeActiveSheetFromWindowActivation (Workbook workbook, Window window, string eventName)
		{
			string workbookKey = GetWorkbookKey (workbook);
			if (string.IsNullOrWhiteSpace (workbookKey) || _activeSheetSynchronizedWorkbookKeys.Contains (workbookKey)) {
				return;
			}
			if (window == null) {
				_logger.Info ("Accounting workbook active sheet sync skipped. reason=" + (eventName ?? string.Empty) + ", windowMissing=True, workbook=" + workbookKey);
				return;
			}
			Worksheet worksheet = null;
			try {
				worksheet = workbook?.ActiveSheet as Worksheet;
				if (worksheet == null) {
					_logger.Info ("Accounting workbook active sheet sync skipped. reason=" + (eventName ?? string.Empty) + ", activeSheetMissing=True, workbook=" + workbookKey);
					return;
				}
				if (HandleActivatedWorksheet (worksheet, eventName, window)) {
					_activeSheetSynchronizedWorkbookKeys.Add (workbookKey);
				}
			} catch (Exception exception) {
				_logger.Error ("Accounting workbook active sheet sync failed.", exception);
			} finally {
				CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (worksheet);
			}
		}

		private bool HandleActivatedWorksheet (Worksheet worksheet, string eventName, Window activatedWindow)
		{
			Workbook workbook = null;
			try {
				workbook = worksheet.Parent as Workbook;
				if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
					return false;
				}
				if (IsCutCopyInProgress (workbook)) {
					_logger.Info ("Accounting workbook sheet activation UI sync skipped because cut/copy mode is active. reason=" + (eventName ?? string.Empty) + ", workbook=" + GetWorkbookKey (workbook));
					return false;
				}
				EnsureWorkbookInitialized (workbook);
				string text = worksheet.CodeName ?? string.Empty;
				if (ShouldSelectTopLeftCell (text)) {
					_accountingWorkbookService.ActivateCell (workbook, text, "A1");
				}
				if (!string.Equals (text, "会計依頼書", StringComparison.OrdinalIgnoreCase)) {
					_accountingWorkbookService.ClearAccountingImportTargetHighlight (workbook);
				}
				_accountingPaymentHistoryImportService.HandleSheetActivated (workbook, text);
				_accountingFormHelperService.HandleSheetActivated (workbook, activatedWindow ?? ResolveActiveSheetWindow (workbook), text);
				return true;
			} catch (Exception exception) {
				_logger.Error ("Accounting workbook sheet activation handling failed.", exception);
				return false;
			}
		}

		internal void RemoveWorkbookState (Workbook workbook)
		{
			string workbookKey = GetWorkbookKey (workbook);
			if (!string.IsNullOrWhiteSpace (workbookKey)) {
				_initializedWorkbookKeys.Remove (workbookKey);
				_activeSheetSynchronizedWorkbookKeys.Remove (workbookKey);
			}
		}

		internal void HandleWorkbookBeforeClose (Workbook workbook)
		{
			if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				return;
			}
			try {
				_accountingPaymentHistoryImportService.HandleWorkbookBeforeClose (workbook);
				_accountingFormHelperService.HandleWorkbookBeforeClose (workbook);
			} catch (Exception exception) {
				_logger.Error ("Accounting workbook before-close handling failed.", exception);
			}
		}

		internal bool TryCancelWorkbookBeforeCloseForActiveAccountingForm (Workbook workbook)
		{
			if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				return false;
			}
			try {
				bool canceled = _accountingFormHelperService.TryCancelWorkbookCloseForActiveAccountingForm (workbook);
				if (canceled) {
					_logger.Info ("Accounting workbook before-close canceled by active accounting form guard.");
				}
				return canceled;
			} catch (Exception exception) {
				_logger.Error ("Accounting workbook form close guard failed.", exception);
				return false;
			}
		}

		private void EnsureWorkbookInitialized (Workbook workbook)
		{
			string workbookKey = GetWorkbookKey (workbook);
			if (string.IsNullOrWhiteSpace (workbookKey) || _initializedWorkbookKeys.Contains (workbookKey)) {
				return;
			}
			using (_accountingWorkbookService.BeginInitializationScope ()) {
				foreach (string allManagedSheetName in GetAllManagedSheetNames ()) {
					_accountingWorkbookService.ProtectSheetUiOnly (workbook, allManagedSheetName);
				}
				_accountingWorkbookService.EnsureNumberFormatLocal (workbook, "見積書", "F15:F33", "\\#,##0;\\-#,##0");
				_accountingWorkbookService.EnsureNumberFormatLocal (workbook, "請求書", "F15:F33", "\\#,##0;\\-#,##0");
				_accountingWorkbookService.EnsureNumberFormatLocal (workbook, "領収書", "F15:F33", "\\#,##0;\\-#,##0");
				_accountingWorkbookService.EnsureNumberFormatLocal (workbook, "会計依頼書", "F15:F33", "\\#,##0;\\-#,##0");
			}
			_initializedWorkbookKeys.Add (workbookKey);
			_logger.Info ("Accounting workbook lifecycle initialization completed. workbook=" + workbookKey);
		}

		private static bool ShouldSelectTopLeftCell (string activeSheetCodeName)
		{
			return string.Equals (activeSheetCodeName, "見積書", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "請求書", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "領収書", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "分割払い予定表", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "お支払い履歴", StringComparison.OrdinalIgnoreCase);
		}

		private static IEnumerable<string> GetAllManagedSheetNames ()
		{
			yield return "見積書";
			yield return "請求書";
			yield return "領収書";
			yield return "会計依頼書";
			yield return "分割払い予定表";
			yield return "お支払い履歴";
			yield return "リカバリ";
			yield return "引数";
		}

		private static bool IsCutCopyInProgress (Workbook workbook)
		{
			if (workbook == null || workbook.Application == null) {
				return false;
			}
			XlCutCopyMode cutCopyMode = workbook.Application.CutCopyMode;
			return cutCopyMode == XlCutCopyMode.xlCopy || cutCopyMode == XlCutCopyMode.xlCut;
		}

		private static Window ResolveActiveSheetWindow (Workbook workbook)
		{
			if (workbook == null) {
				return null;
			}
			try {
				Application application = workbook.Application;
				Window activeWindow = (application == null) ? null : application.ActiveWindow;
				if (activeWindow != null) {
					return activeWindow;
				}
			} catch {
			}
			return GetFirstVisibleWindow (workbook);
		}

		private static Window GetFirstVisibleWindow (Workbook workbook)
		{
			Windows windows = null;
			try {
				windows = workbook?.Windows;
				int windowCount = windows == null ? 0 : windows.Count;
				for (int index = 1; index <= windowCount; index++) {
					Window window = null;
					try {
						window = windows [index];
						if (window != null && window.Visible) {
							return window;
						}
					} finally {
						if (window != null && !window.Visible) {
							CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (window);
						}
					}
				}
			} catch {
				return null;
			} finally {
				CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (windows);
			}
			return null;
		}

		private static string GetWorkbookKey (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			string text = workbook.FullName ?? string.Empty;
			return string.IsNullOrWhiteSpace (text) ? (workbook.Name ?? string.Empty) : text;
		}

	}
}
