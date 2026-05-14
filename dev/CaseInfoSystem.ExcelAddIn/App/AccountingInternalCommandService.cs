using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingInternalCommandService
	{
		private readonly NavigationService _navigationService;

		private readonly AccountingPaymentHistoryImportService _accountingPaymentHistoryImportService;

		private readonly AccountingFormHelperService _accountingFormHelperService;

		private readonly AccountingSaveAsService _accountingSaveAsService;

		private readonly Logger _logger;

		internal AccountingInternalCommandService (NavigationService navigationService, AccountingPaymentHistoryImportService accountingPaymentHistoryImportService, AccountingFormHelperService accountingFormHelperService, AccountingSaveAsService accountingSaveAsService, Logger logger)
		{
			_navigationService = navigationService ?? throw new ArgumentNullException ("navigationService");
			_accountingPaymentHistoryImportService = accountingPaymentHistoryImportService ?? throw new ArgumentNullException ("accountingPaymentHistoryImportService");
			_accountingFormHelperService = accountingFormHelperService ?? throw new ArgumentNullException ("accountingFormHelperService");
			_accountingSaveAsService = accountingSaveAsService ?? throw new ArgumentNullException ("accountingSaveAsService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (WorkbookContext context, string actionId)
		{
			if (context == null) {
				throw new ArgumentNullException ("context");
			}
			if (!string.IsNullOrWhiteSpace (actionId)) {
				_logger.Info ("Accounting pane action requested. actionId=" + actionId + ", workbook=" + (context.WorkbookFullName ?? string.Empty));
				switch (actionId) {
				case "import-payment-history-to-request":
					_accountingPaymentHistoryImportService.Execute (context);
					break;
				case AccountingNavigationActionIds.SetInstallmentScheduleIssueDate:
				case AccountingNavigationActionIds.ResetInstallmentSchedule:
				case AccountingNavigationActionIds.SetPaymentHistoryIssueDate:
				case AccountingNavigationActionIds.DeleteSelectedPaymentHistoryRows:
				case AccountingNavigationActionIds.ResetPaymentHistory:
				case "set-issue-date":
				case "set-issue-date-and-due-date":
				case "open-payment-history-input":
				case "open-installment-schedule-input":
				case "open-reverse-tool":
					_accountingFormHelperService.Execute (context, actionId);
					break;
				case "switch-to-estimate-sheet":
					SwitchSheet (context, "見積書");
					break;
				case "switch-to-invoice-sheet":
					SwitchSheet (context, "請求書");
					break;
				case "switch-to-receipt-sheet":
					SwitchSheet (context, "領収書");
					break;
				case "switch-to-accounting-request-sheet":
					SwitchSheet (context, "会計依頼書");
					break;
				case "switch-to-installment-sheet":
					SwitchSheet (context, "分割払い予定表");
					break;
				case "switch-to-payment-history-sheet":
					SwitchSheet (context, "お支払い履歴");
					break;
				case "show-save-as-prompt":
					_accountingSaveAsService.Execute (context);
					break;
				}
			}
		}

		private void SwitchSheet (WorkbookContext context, string targetSheetCodeName)
		{
			if (!_navigationService.TryNavigateToSheet (context.Workbook, targetSheetCodeName, "AccountingPaneSwitch")) {
				throw new InvalidOperationException ("遷移先シートを表示できませんでした。sheetCodeName=" + targetSheetCodeName);
			}
		}

		private static void ShowPendingMessage (string operationName)
		{
			UserErrorService.ShowOkNotification (operationName + " の実行フローは今回の対象外として後回しにしています。", "案件情報System", MessageBoxIcon.Asterisk);
		}
	}
}
