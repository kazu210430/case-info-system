using System;
using System.Globalization;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingInstallmentScheduleCommandService
	{
		private const string SheetName = "分割払い予定表";

		private const string InstallmentIssueDateRangeName = "発行日";

		private const string InstallmentBillingAmountRangeName = "請求額";

		private const string InstallmentExpenseAmountRangeName = "実費等総額";

		private const string InstallmentWithholdingRangeName = "源泉処理";

		private const string InstallmentFirstDueDateRangeName = "第1回期限";

		private const string InstallmentAmountRangeName = "分割金";

		private const string InstallmentChangeRoundRangeName = "変更回";

		private const string InstallmentChangedAmountRangeName = "変更後分割金";

		private const string InstallmentDepositAmountRangeName = "お預かり金額";

		private const string InstallmentPaymentTotalCellAddress = "I9";

		private const string InstallmentMarkerCellAddress = "B12";

		private const string InstallmentRoundStartCellAddress = "A13";

		private const string InvoiceBillingSubtotalCellAddress = "F23";

		private const string InvoiceExpenseCellAddress = "F25";

		private const string InvoiceWithholdingFlagCellAddress = "Y24";

		private const string InvoiceFirstDueDateCellAddress = "G10";

		private const string InvoiceDepositAmountCellAddress = "F29";

		private const string InstallmentIssueDateSourceCellAddress = "J1";

		private const string InstallmentResetMarkerText = "※";

		private const string InstallmentIssueDatePlaceholderText = "発行日";

		private const string DateDisplayFormat = "yyyy/MM/dd";

		private const int InstallmentResetStartRow = 13;

		private const int InstallmentResetEndRow = 73;

		private const int InstallmentDefaultPrintLastRow = 13;

		private const int InstallmentDefaultPrintLastColumn = 9;

		private const int InstallmentMaximumRounds = 60;

		private const int InstallmentFirstRow = 13;

		private const int InstallmentLastColumn = 9;

		private const string DepositAppliedText = "(充当済み)";

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly UserErrorService _userErrorService;

		private readonly Logger _logger;

		internal AccountingInstallmentScheduleCommandService (AccountingWorkbookService accountingWorkbookService, UserErrorService userErrorService, Logger logger)
		{
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_userErrorService = userErrorService ?? throw new ArgumentNullException ("userErrorService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal AccountingInstallmentScheduleFormState LoadFormState (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			SyncInstallmentHeaderFromInvoice (workbook);
			double num = ReadDoubleSafe (workbook, "請求書", "F23");
			double num2 = ReadDoubleSafe (workbook, "請求書", "F25");
			bool flag = ReadBooleanSafe (workbook, "請求書", "Y24");
			return new AccountingInstallmentScheduleFormState {
				BillingAmountText = FormatAmount (num + num2),
				ExpenseAmountText = FormatAmount (num2),
				WithholdingText = (flag ? "する" : "しない"),
				FirstDueDateText = ReadFormattedDateSafe (workbook, "請求書", "G10"),
				DepositAmountText = FormatAmount (ReadDoubleSafe (workbook, "請求書", "F29")),
				InstallmentAmountText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, "分割払い予定表", "分割金"),
				ChangeRoundText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, "分割払い予定表", "変更回"),
				ChangedInstallmentAmountText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, "分割払い予定表", "変更後分割金")
			};
		}

		internal AccountingInstallmentScheduleFormState CreateSchedule (Workbook workbook, AccountingInstallmentScheduleCreateRequest request)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			if (!TryValidateCreateRequest (request, out var firstDueDate, out var billingAmount, out var expenseAmount, out var depositAmount, out var installmentAmount)) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "分割払い予定表");
				SyncInstallmentHeaderFromInvoice (workbook);
				WriteCreateBaseValues (workbook, request, firstDueDate, billingAmount, expenseAmount, depositAmount, installmentAmount);
				int num = ((depositAmount != 0.0) ? BuildScheduleWithDeposit (workbook, firstDueDate, depositAmount, installmentAmount) : BuildScheduleWithoutDeposit (workbook, firstDueDate, installmentAmount));
				RefreshPrintArea (workbook);
				WritePaymentTotal (workbook);
				_logger.Info ("Installment schedule created. lastInputRow=" + num.ToString (CultureInfo.InvariantCulture));
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.CreateSchedule", exception);
			} finally {
				try {
					_accountingWorkbookService.ProtectSheetUiOnly (workbook, "分割払い予定表");
				} catch (Exception exception2) {
					_logger.Error ("Installment schedule reprotect failed after create.", exception2);
				}
			}
			return LoadFormState (workbook);
		}

		internal AccountingInstallmentScheduleFormState ApplyChange (Workbook workbook, AccountingInstallmentScheduleChangeRequest request)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			if (!TryValidateChangeRequest (workbook, request, out var scheduleStartFlag, out var startRow, out var changedInstallmentAmount, out var billingAmount, out var expenseAmount)) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "分割払い予定表");
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "変更回", ParseAmount (request.ChangeRoundText));
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "請求額", billingAmount);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "実費等総額", expenseAmount);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "源泉処理", request.WithholdingText ?? string.Empty);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "変更後分割金", changedInstallmentAmount);
				int num = BuildChangedSchedule (workbook, scheduleStartFlag, startRow, changedInstallmentAmount);
				RefreshPrintArea (workbook);
				WritePaymentTotal (workbook);
				_logger.Info ("Installment schedule change applied. lastInputRow=" + num.ToString (CultureInfo.InvariantCulture));
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.ApplyChange", exception);
			} finally {
				try {
					_accountingWorkbookService.ProtectSheetUiOnly (workbook, "分割払い予定表");
				} catch (Exception exception2) {
					_logger.Error ("Installment schedule reprotect failed after change.", exception2);
				}
			}
			return LoadFormState (workbook);
		}

		internal AccountingInstallmentScheduleFormState ApplyIssueDate (Workbook workbook)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, "分割払い予定表", "J1");
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "発行日", value);
				_logger.Info ("Installment schedule issue date applied.");
				return LoadFormState (workbook);
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.ApplyIssueDate", exception);
				return LoadFormState (workbook);
			}
		}

		internal AccountingInstallmentScheduleFormState Reset (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			DialogResult dialogResult = MessageBox.Show ("分割払い予定表は全てクリアされます。よろしいですか？", "案件情報System", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
			if (dialogResult != DialogResult.OK) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "分割払い予定表");
				SyncInstallmentHeaderFromInvoice (workbook);
				_accountingWorkbookService.WriteCell (workbook, "分割払い予定表", "B12", "※");
				_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "A13", 1);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "発行日", "発行日");
				_accountingWorkbookService.ClearNamedRangeMergeAreaContents (workbook, "分割払い予定表", "請求額");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "分割金");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "第1回期限");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "実費等総額");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "源泉処理");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "変更回");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "変更後分割金");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "分割払い予定表", "お預かり金額");
				ClearRowsFrom (13, workbook);
				_accountingWorkbookService.ClearRangeContents (workbook, "分割払い予定表", "I9");
				_accountingWorkbookService.SetPrintAreaByBounds (workbook, "分割払い予定表", 13, 9);
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.Reset", exception);
			} finally {
				try {
					_accountingWorkbookService.ProtectSheetUiOnly (workbook, "分割払い予定表");
				} catch (Exception exception2) {
					_logger.Error ("Installment schedule reprotect failed.", exception2);
				}
			}
			return LoadFormState (workbook);
		}

		private bool TryValidateCreateRequest (AccountingInstallmentScheduleCreateRequest request, out DateTime firstDueDate, out double billingAmount, out double expenseAmount, out double depositAmount, out double installmentAmount)
		{
			firstDueDate = DateTime.MinValue;
			billingAmount = ParseAmount (request.BillingAmountText);
			expenseAmount = ParseAmount (request.ExpenseAmountText);
			depositAmount = ParseAmount (request.DepositAmountText);
			installmentAmount = ParseAmount (request.InstallmentAmountText);
			if (billingAmount == 0.0) {
				MessageBox.Show ("請求額が0円です。請求書を作成してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if (!DateTime.TryParse (request.FirstDueDateText, CultureInfo.InvariantCulture, DateTimeStyles.None, out firstDueDate) && !DateTime.TryParse (request.FirstDueDateText, CultureInfo.CurrentCulture, DateTimeStyles.None, out firstDueDate)) {
				MessageBox.Show ("期限（1回目）欄に日付が正しく入力されていません\r\n請求書で期限を入力してください", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if (string.IsNullOrWhiteSpace (request.InstallmentAmountText) || installmentAmount <= 0.0) {
				MessageBox.Show ("分割払い額が未入力です", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if (billingAmount / installmentAmount > 60.0) {
				MessageBox.Show ("分割回数が60回を超えてしまいます", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			return true;
		}

		private bool TryValidateChangeRequest (Workbook workbook, AccountingInstallmentScheduleChangeRequest request, out int scheduleStartFlag, out int startRow, out double changedInstallmentAmount, out double billingAmount, out double expenseAmount)
		{
			scheduleStartFlag = Convert.ToInt32 (ReadCellDouble (workbook, "A13"), CultureInfo.InvariantCulture);
			startRow = 0;
			changedInstallmentAmount = ParseAmount (request.ChangedInstallmentAmountText);
			billingAmount = ParseAmount (request.BillingAmountText);
			expenseAmount = ParseAmount (request.ExpenseAmountText);
			double num = ParseAmount (request.ChangeRoundText);
			if (billingAmount == 0.0) {
				MessageBox.Show ("請求額が0円です。請求書を作成してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if (string.IsNullOrWhiteSpace (request.ChangeRoundText)) {
				MessageBox.Show ("変更する回が未入力です。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if (num < 2.0 || num > 60.0) {
				MessageBox.Show ("途中変更する回は2～60の数値を入力してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if (string.IsNullOrWhiteSpace (request.ChangedInstallmentAmountText) || changedInstallmentAmount <= 0.0) {
				MessageBox.Show ("変更後の分割払い額が未入力です。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			int num2 = Convert.ToInt32 (num, CultureInfo.InvariantCulture) + 13;
			int num3 = Convert.ToInt32 (num, CultureInfo.InvariantCulture) + 12;
			startRow = ((scheduleStartFlag == 0) ? num2 : num3);
			if ((scheduleStartFlag == 0 && ReadCellDouble (workbook, "I" + (num2 - 1).ToString (CultureInfo.InvariantCulture)) <= 0.0) || (scheduleStartFlag == 1 && ReadCellDouble (workbook, "I" + (num3 - 1).ToString (CultureInfo.InvariantCulture)) <= 0.0)) {
				MessageBox.Show ("変更する回にはすでに完済しています。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			if ((scheduleStartFlag == 0 && ReadCellDouble (workbook, "I" + (num2 - 1).ToString (CultureInfo.InvariantCulture)) / changedInstallmentAmount >= 60.0 - (double)(num2 - 13) + 1.0) || (scheduleStartFlag == 1 && ReadCellDouble (workbook, "I" + (num3 - 1).ToString (CultureInfo.InvariantCulture)) / changedInstallmentAmount >= 60.0 - (double)(num3 - 12) + 1.0)) {
				MessageBox.Show ("分割回数が60回を超えてしまいます。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return false;
			}
			return true;
		}

		private void WriteCreateBaseValues (Workbook workbook, AccountingInstallmentScheduleCreateRequest request, DateTime firstDueDate, double billingAmount, double expenseAmount, double depositAmount, double installmentAmount)
		{
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "請求額", billingAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "実費等総額", expenseAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "源泉処理", request.WithholdingText ?? string.Empty);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "第1回期限", firstDueDate);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "分割金", installmentAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "分割払い予定表", "お預かり金額", depositAmount);
		}

		private int BuildScheduleWithDeposit (Workbook workbook, DateTime firstDueDate, double depositAmount, double installmentAmount)
		{
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "A13", 0);
			_accountingWorkbookService.WriteCell (workbook, "分割払い予定表", "B13", "(充当済み)");
			_accountingWorkbookService.SetHorizontalAlignmentCenter (workbook, "分割払い予定表", "B13");
			double val = ReadCellDouble (workbook, "J12");
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "H13", Math.Min (depositAmount, val));
			ExecuteGoalSeekAndRound (workbook, 13, "C", "D", depositAmount);
			int num = 14;
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "B14", firstDueDate);
			ApplyExpenseForRow (workbook, num, installmentAmount);
			ExecuteGoalSeekAndRound (workbook, num, "C", "D", installmentAmount);
			return CompleteTrailingRowsAfterLoop (workbook, num, installmentAmount);
		}

		private int BuildScheduleWithoutDeposit (Workbook workbook, DateTime firstDueDate, double installmentAmount)
		{
			int num = 13;
			WriteDueDate (workbook, num, firstDueDate);
			ApplyExpenseForRow (workbook, num, installmentAmount);
			ExecuteGoalSeekAndRound (workbook, num, "C", "D", installmentAmount);
			return CompleteTrailingRowsAfterLoop (workbook, num, installmentAmount);
		}

		private int BuildChangedSchedule (Workbook workbook, int scheduleStartFlag, int startRow, double changedInstallmentAmount)
		{
			WriteDueDate (workbook, startRow, DateTime.MinValue);
			ApplyExpenseForChangedRow (workbook, startRow, changedInstallmentAmount);
			ExecuteGoalSeekAndRound (workbook, startRow, "C", "D", changedInstallmentAmount);
			int num = startRow;
			while (ReadCellDouble (workbook, "I" + num.ToString (CultureInfo.InvariantCulture)) > changedInstallmentAmount) {
				int num2 = num + 1;
				WriteDueDate (workbook, num2, DateTime.MinValue);
				ApplyExpenseForChangedRow (workbook, num2, changedInstallmentAmount);
				ExecuteGoalSeekAndRound (workbook, num2, "C", "D", changedInstallmentAmount);
				num = num2;
			}
			if (ReadCellDouble (workbook, "I" + num.ToString (CultureInfo.InvariantCulture)) < 0.0) {
				AdjustRow (workbook, num);
				ClearRowsFrom (num + 1, workbook);
				return num;
			}
			int num3 = num + 1;
			WriteDueDate (workbook, num3, DateTime.MinValue);
			ApplyExpenseForChangedRow (workbook, num3, changedInstallmentAmount);
			AdjustRow (workbook, num3);
			ClearRowsFrom (num3 + 1, workbook);
			return num3;
		}

		private int CompleteTrailingRowsAfterLoop (Workbook workbook, int currentRow, double installmentAmount)
		{
			while (ReadCellDouble (workbook, "I" + currentRow.ToString (CultureInfo.InvariantCulture)) > installmentAmount) {
				int num = currentRow + 1;
				WriteDueDate (workbook, num, DateTime.MinValue);
				ApplyExpenseForRow (workbook, num, installmentAmount);
				ExecuteGoalSeekAndRound (workbook, num, "C", "D", installmentAmount);
				currentRow = num;
			}
			if (ReadCellDouble (workbook, "I" + currentRow.ToString (CultureInfo.InvariantCulture)) < 0.0) {
				AdjustRow (workbook, currentRow);
				ClearRowsFrom (currentRow + 1, workbook);
				return currentRow;
			}
			int num2 = currentRow + 1;
			WriteDueDate (workbook, num2, DateTime.MinValue);
			ApplyExpenseForRow (workbook, num2, installmentAmount);
			AdjustRow (workbook, num2);
			ClearRowsFrom (num2 + 1, workbook);
			return num2;
		}

		private void WriteDueDate (Workbook workbook, int row, DateTime firstDueDate)
		{
			if (row == 13) {
				_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "B13", firstDueDate);
				return;
			}
			DateTime dateTime = ReadCellDate (workbook, "B" + (row - 1).ToString (CultureInfo.InvariantCulture));
			DateTime dateTime2 = new DateTime (dateTime.Year, dateTime.Month, 1).AddMonths (2).AddDays (-1.0);
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "B" + row.ToString (CultureInfo.InvariantCulture), dateTime2);
		}

		private void ApplyExpenseForRow (Workbook workbook, int row, double installmentAmount)
		{
			double num = ReadCellDouble (workbook, "J" + (row - 1).ToString (CultureInfo.InvariantCulture));
			double num2 = ((installmentAmount >= num) ? num : installmentAmount);
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "H" + row.ToString (CultureInfo.InvariantCulture), num2);
		}

		private void ApplyExpenseForChangedRow (Workbook workbook, int row, double changedInstallmentAmount)
		{
			double num = ReadCellDouble (workbook, "J" + (row - 1).ToString (CultureInfo.InvariantCulture));
			double num2 = ((changedInstallmentAmount >= num) ? num : changedInstallmentAmount);
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "H" + row.ToString (CultureInfo.InvariantCulture), num2);
		}

		private void ExecuteGoalSeekAndRound (Workbook workbook, int row, string formulaColumn, string changingColumn, double goalValue)
		{
			string text = row.ToString (CultureInfo.InvariantCulture);
			string formulaCellAddress = formulaColumn + text;
			string text2 = changingColumn + text;
			_accountingWorkbookService.ExecuteGoalSeek (workbook, "分割払い予定表", formulaCellAddress, text2, goalValue);
			double value = ReadCellDouble (workbook, text2);
			double num = Math.Round (value, 0, MidpointRounding.AwayFromZero);
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", text2, num);
		}

		private void AdjustRow (Workbook workbook, int row)
		{
			double goalValue = ReadCellDouble (workbook, "I" + (row - 1).ToString (CultureInfo.InvariantCulture)) - ReadCellDouble (workbook, "H" + row.ToString (CultureInfo.InvariantCulture));
			ExecuteGoalSeekAndRound (workbook, row, "F", "D", goalValue);
		}

		private void ClearRowsFrom (int startRow, Workbook workbook)
		{
			if (startRow <= 73) {
				string text = startRow.ToString (CultureInfo.InvariantCulture);
				string text2 = 73.ToString (CultureInfo.InvariantCulture);
				_accountingWorkbookService.ClearRangeContents (workbook, "分割払い予定表", "B" + text + ":B" + text2);
				_accountingWorkbookService.ClearRangeContents (workbook, "分割払い予定表", "D" + text + ":D" + text2);
				_accountingWorkbookService.ClearRangeContents (workbook, "分割払い予定表", "H" + text + ":H" + text2);
			}
		}

		private void RefreshPrintArea (Workbook workbook)
		{
			int lastUsedRowInColumn = _accountingWorkbookService.GetLastUsedRowInColumn (workbook, "分割払い予定表", "D");
			if (lastUsedRowInColumn >= 1) {
				_accountingWorkbookService.SetPrintAreaByBounds (workbook, "分割払い予定表", lastUsedRowInColumn, 9);
			}
		}

		private void WritePaymentTotal (Workbook workbook)
		{
			int lastUsedRowInColumn = _accountingWorkbookService.GetLastUsedRowInColumn (workbook, "分割払い予定表", "C");
			double num = 0.0;
			for (int i = 13; i <= lastUsedRowInColumn; i++) {
				num += ReadCellDouble (workbook, "C" + i.ToString (CultureInfo.InvariantCulture));
			}
			_accountingWorkbookService.WriteCellValue (workbook, "分割払い予定表", "I9", num);
		}

		private static double ParseAmount (string text)
		{
			string text2 = (text ?? string.Empty).Replace (",", string.Empty).Trim ();
			if (text2.Length == 0) {
				return 0.0;
			}
			double result;
			return double.TryParse (text2, NumberStyles.Number, CultureInfo.InvariantCulture, out result) ? result : 0.0;
		}

		private double ReadCellDouble (Workbook workbook, string address)
		{
			object obj = _accountingWorkbookService.ReadCellValue (workbook, "分割払い予定表", address);
			try {
				return (obj == null) ? 0.0 : Convert.ToDouble (obj, CultureInfo.InvariantCulture);
			} catch {
				return 0.0;
			}
		}

		private DateTime ReadCellDate (Workbook workbook, string address)
		{
			return _accountingWorkbookService.ReadDateCell (workbook, "分割払い予定表", address);
		}

		private void SyncInstallmentHeaderFromInvoice (Workbook workbook)
		{
			_accountingWorkbookService.CopyValueRange (workbook, "請求書", "A3:A4", "分割払い予定表", "A3:A4");
		}

		private double ReadDoubleSafe (Workbook workbook, string sheetName, string address)
		{
			object obj = _accountingWorkbookService.ReadCellValue (workbook, sheetName, address);
			try {
				return (obj == null) ? 0.0 : Convert.ToDouble (obj, CultureInfo.InvariantCulture);
			} catch {
				return 0.0;
			}
		}

		private bool ReadBooleanSafe (Workbook workbook, string sheetName, string address)
		{
			object obj = _accountingWorkbookService.ReadCellValue (workbook, sheetName, address);
			try {
				return obj != null && Convert.ToBoolean (obj, CultureInfo.InvariantCulture);
			} catch {
				return false;
			}
		}

		private static string FormatAmount (double value)
		{
			return value.ToString ("#,##0", CultureInfo.InvariantCulture);
		}

		private string ReadFormattedDateSafe (Workbook workbook, string sheetName, string address)
		{
			try {
				return _accountingWorkbookService.ReadDateCell (workbook, sheetName, address).ToString ("yyyy/MM/dd", CultureInfo.InvariantCulture);
			} catch {
				return _accountingWorkbookService.ReadDisplayText (workbook, sheetName, address);
			}
		}
	}
}
