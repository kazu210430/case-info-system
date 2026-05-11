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

		private const string PaymentHistoryChargedMarker = "(充当済み)";

		private const string PaymentHistoryDefaultSortRangeAddress = "B13:H72";

		private const string PaymentHistoryDefaultSortKeyAddress = "B13";

		private const string PaymentHistoryChargedSortRangeAddress = "B14:H72";

		private const string PaymentHistoryChargedSortKeyAddress = "B14";

		private readonly WorkbookRoleResolver _workbookRoleResolver;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly AccountingFormHelperService _accountingFormHelperService;

		private readonly AccountingPaymentHistoryImportService _accountingPaymentHistoryImportService;

		private readonly PostCloseFollowUpScheduler _postCloseFollowUpScheduler;

		private readonly Logger _logger;

		private readonly HashSet<string> _initializedWorkbookKeys;

		internal AccountingWorkbookLifecycleService (WorkbookRoleResolver workbookRoleResolver, AccountingWorkbookService accountingWorkbookService, AccountingFormHelperService accountingFormHelperService, AccountingPaymentHistoryImportService accountingPaymentHistoryImportService, Logger logger)
			: this (workbookRoleResolver, accountingWorkbookService, accountingFormHelperService, accountingPaymentHistoryImportService, null, logger)
		{
		}

		internal AccountingWorkbookLifecycleService (WorkbookRoleResolver workbookRoleResolver, AccountingWorkbookService accountingWorkbookService, AccountingFormHelperService accountingFormHelperService, AccountingPaymentHistoryImportService accountingPaymentHistoryImportService, PostCloseFollowUpScheduler postCloseFollowUpScheduler, Logger logger)
		{
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_accountingFormHelperService = accountingFormHelperService ?? throw new ArgumentNullException ("accountingFormHelperService");
			_accountingPaymentHistoryImportService = accountingPaymentHistoryImportService ?? throw new ArgumentNullException ("accountingPaymentHistoryImportService");
			_postCloseFollowUpScheduler = postCloseFollowUpScheduler;
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_initializedWorkbookKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
		}

		internal void HandleWorkbookOpenedOrActivated (Workbook workbook)
		{
			if (_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				EnsureWorkbookInitialized (workbook);
			}
		}

		internal void HandleSheetActivated (object sheetObject)
		{
			if (!(sheetObject is Worksheet worksheet)) {
				return;
			}
			Workbook workbook = null;
			try {
				workbook = worksheet.Parent as Workbook;
				if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
					return;
				}
				if (IsCutCopyInProgress (workbook)) {
					_logger.Info ("Accounting workbook sheet activation UI sync skipped because cut/copy mode is active. workbook=" + GetWorkbookKey (workbook));
					return;
				}
				EnsureWorkbookInitialized (workbook);
				string text = worksheet.CodeName ?? string.Empty;
				if (ShouldSelectTopLeftCell (text)) {
					_accountingWorkbookService.ActivateCell (workbook, text, "A1");
				}
				if (string.Equals (text, "お支払い履歴", StringComparison.OrdinalIgnoreCase)) {
					ApplyPaymentHistorySort (workbook);
				}
				if (!string.Equals (text, "会計依頼書", StringComparison.OrdinalIgnoreCase)) {
					_accountingWorkbookService.ClearAccountingImportTargetHighlight (workbook);
				}
				_accountingPaymentHistoryImportService.HandleSheetActivated (workbook, text);
				Application application = workbook.Application;
				_accountingFormHelperService.HandleSheetActivated (workbook, (application == null) ? null : application.ActiveWindow, text);
			} catch (Exception exception) {
				_logger.Error ("Accounting workbook sheet activation handling failed.", exception);
			}
		}

		internal void RemoveWorkbookState (Workbook workbook)
		{
			string workbookKey = GetWorkbookKey (workbook);
			if (!string.IsNullOrWhiteSpace (workbookKey)) {
				_initializedWorkbookKeys.Remove (workbookKey);
			}
		}

		internal void HandleWorkbookBeforeClose (Workbook workbook)
		{
			if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				return;
			}
			string workbookKey = GetWorkbookKey (workbook);
			try {
				_accountingPaymentHistoryImportService.HandleWorkbookBeforeClose (workbook);
				_accountingFormHelperService.HandleWorkbookBeforeClose (workbook);
			} catch (Exception exception) {
				_logger.Error ("Accounting workbook before-close handling failed.", exception);
			}
			SchedulePostCloseFollowUp (workbookKey, GetWorkbookFolderPath (workbook));
		}

		private void SchedulePostCloseFollowUp (string workbookKey, string folderPath)
		{
			if (_postCloseFollowUpScheduler == null || string.IsNullOrWhiteSpace (workbookKey)) {
				return;
			}
			_logger.Info ("Accounting workbook post-close follow-up scheduled. workbook=" + workbookKey);
			_postCloseFollowUpScheduler.ScheduleManagedWorkbookClose (workbookKey, folderPath, ManagedWorkbookCloseMarkerKind.AccountingClose);
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

		private void ApplyPaymentHistorySort (Workbook workbook)
		{
			string text = _accountingWorkbookService.ReadDisplayText (workbook, "お支払い履歴", "B13");
			bool flag = string.Equals (text, "(充当済み)", StringComparison.OrdinalIgnoreCase);
			string text2 = (flag ? "B14:H72" : "B13:H72");
			string text3 = (flag ? "B14" : "B13");
			_accountingWorkbookService.SortRangeAscending (workbook, "お支払い履歴", text2, text3);
			_logger.Info ("Accounting payment history sorted on sheet activation. firstMarker=" + (text ?? string.Empty) + ", range=" + text2 + ", key=" + text3);
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

		private static string GetWorkbookKey (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			string text = workbook.FullName ?? string.Empty;
			return string.IsNullOrWhiteSpace (text) ? (workbook.Name ?? string.Empty) : text;
		}

		private static string GetWorkbookFolderPath (Workbook workbook)
		{
			try {
				return workbook == null ? string.Empty : workbook.Path ?? string.Empty;
			} catch {
				return string.Empty;
			}
		}
	}
}
