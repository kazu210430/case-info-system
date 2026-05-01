using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingPaymentHistoryCommandService
	{
		private sealed class PaymentHistoryEditableRow
		{
			internal object DateValue { get; }

			internal object TargetAmountValue { get; }

			internal object ExpenseAmountValue { get; }

			internal DateTime SortDate { get; }

			internal PaymentHistoryEditableRow (object dateValue, object targetAmountValue, object expenseAmountValue, DateTime sortDate)
			{
				DateValue = dateValue;
				TargetAmountValue = targetAmountValue;
				ExpenseAmountValue = expenseAmountValue;
				SortDate = sortDate;
			}
		}

		private const string SheetName = "お支払い履歴";

		private const string InvoiceSheetName = "請求書";

		private const string PaymentHistoryIssueDateRangeName = "発行日";

		private const string PaymentHistoryBillingAmountRangeName = "請求額";

		private const string PaymentHistoryReceiptDateRangeName = "領収日";

		private const string PaymentHistoryAmountRangeName = "分割金";

		private const string PaymentHistoryExpenseAmountRangeName = "実費等総額";

		private const string PaymentHistoryWithholdingRangeName = "源泉処理";

		private const string PaymentHistoryDepositAmountRangeName = "お預かり金額";

		private const string PaymentHistoryDepositProcessedRangeName = "お預かり金額処理";

		private const string PaymentTotalCellAddress = "I9";

		private const string MarkerCellAddress = "B12";

		private const string FirstRoundMarkerCellAddress = "A13";

		private const string FirstReceiptDateCellAddress = "B13";

		private const string InvoiceBillingSubtotalCellAddress = "F23";

		private const string InvoiceExpenseCellAddress = "F25";

		private const string InvoiceWithholdingFlagCellAddress = "Y24";

		private const string InvoiceDepositAmountCellAddress = "F29";

		private const string IssueDateSourceCellAddress = "J1";

		private const string ResetMarkerText = "※";

		private const string IssueDatePlaceholderText = "発行日";

		private const string DateDisplayFormat = "yyyy/MM/dd";

		private const string DepositAppliedText = "(充当済み)";

		private const string DepositProcessedText = "済み";

		private const int FirstHistoryRow = 13;

		private const int LastHistoryRow = 72;

		private const int ClearEndRow = 73;

		private const int PrintLastColumn = 9;

		private const int MaximumRounds = 60;

		private const int PaymentHistoryFirstRoundOffset = 12;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly UserErrorService _userErrorService;

		private readonly Logger _logger;

		internal AccountingPaymentHistoryCommandService (AccountingWorkbookService accountingWorkbookService, UserErrorService userErrorService, Logger logger)
		{
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_userErrorService = userErrorService ?? throw new ArgumentNullException ("userErrorService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal AccountingPaymentHistoryFormState LoadFormState (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			SyncPaymentHistoryHeaderFromInvoice (workbook);
			const string text = "AccountingPaymentHistory.LoadFormState";
			List<string> list = new List<string> ();
			double num = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, "請求書", "F23", "請求額小計", text, list);
			double num2 = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, "請求書", "F25", "実費等総額", text, list);
			double num3 = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, "請求書", "F29", "お預かり金額", text, list);
			bool flag = ReadBooleanSafe (workbook, "請求書", "Y24");
			AccountingPaymentHistoryFormState accountingPaymentHistoryFormState = new AccountingPaymentHistoryFormState {
				BillingAmountText = FormatAmount (num + num2),
				ExpenseAmountText = FormatAmount (num2),
				WithholdingText = (flag ? "する" : "しない"),
				DepositAmountText = FormatAmount (num3),
				ReceiptDateText = ReadFormattedDateFromNamedRangeSafe (workbook, "お支払い履歴", "領収日"),
				ReceiptAmountText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, "お支払い履歴", "分割金"),
				HasNumericReadError = list.Count > 0,
				NumericReadErrorMessage = string.Join (Environment.NewLine, list)
			};
			if (list.Count > 0) {
				ShowLoadFormStateNumericReadWarning (accountingPaymentHistoryFormState.NumericReadErrorMessage);
			}
			return accountingPaymentHistoryFormState;
		}

		internal AccountingPaymentHistoryFormState ApplyIssueDate (Workbook workbook)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, "お支払い履歴", "J1");
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "発行日", value);
				_logger.Info ("Payment history issue date applied.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.ApplyIssueDate", exception);
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState ApplyToday (Workbook workbook)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, "お支払い履歴", "J1");
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "領収日", value);
				_logger.Info ("Payment history today applied from J1.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.ApplyToday", exception);
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState AddHistoryEntry (Workbook workbook, AccountingPaymentHistoryEntryRequest request)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			if (!TryValidateEntryRequest (request, out var receiptDate, out var billingAmount, out var expenseAmount, out var depositAmount, out var receiptAmount)) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "お支払い履歴");
				CompactBlankDateRows (workbook, GetEditableStartRow (workbook));
				WriteBaseValues (workbook, request, receiptDate, billingAmount, expenseAmount, depositAmount, receiptAmount);
				if (depositAmount != 0.0 && !IsDepositProcessed (workbook)) {
					ApplyDepositRow (workbook, depositAmount);
					int nextAppendRow = GetNextAppendRow (workbook);
					WriteReceiptDateRow (workbook, nextAppendRow, receiptDate);
					ApplyExpenseAmount (workbook, nextAppendRow, receiptAmount);
					ExecuteReceiptGoalSeek (workbook, nextAppendRow, receiptAmount);
				} else if (string.IsNullOrWhiteSpace (_accountingWorkbookService.ReadText (workbook, "お支払い履歴", "B13"))) {
					WriteReceiptDateRow (workbook, 13, receiptDate);
					ApplyExpenseAmount (workbook, 13, receiptAmount);
					ExecuteReceiptGoalSeek (workbook, 13, receiptAmount);
				} else {
					int nextAppendRow2 = GetNextAppendRow (workbook);
					WriteReceiptDateRow (workbook, nextAppendRow2, receiptDate);
					ApplyExpenseAmount (workbook, nextAppendRow2, receiptAmount);
					ExecuteReceiptGoalSeek (workbook, nextAppendRow2, receiptAmount);
				}
				int startRow = CompactBlankDateRows (workbook, GetEditableStartRow (workbook));
				ClearRowsFrom (workbook, startRow);
				RefreshPrintArea (workbook);
				WritePaymentTotal (workbook);
				_logger.Info ("Payment history entry added.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.AddHistoryEntry", exception);
			} finally {
				ProtectSheetSafely (workbook, "Payment history add entry reprotect failed.");
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState OutputFutureBalance (Workbook workbook, AccountingPaymentHistoryEntryRequest request)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			if (!TryValidateProjectionRequest (workbook, request, out var billingAmount, out var expenseAmount, out var receiptAmount)) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "お支払い履歴");
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "請求額", billingAmount);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "実費等総額", expenseAmount);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "源泉処理", request.WithholdingText ?? string.Empty);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "分割金", receiptAmount);
				int lastOccupiedHistoryRow = GetLastOccupiedHistoryRow (workbook, GetEditableStartRow (workbook));
				int num = lastOccupiedHistoryRow + 1;
				ApplyExpenseAmount (workbook, num, receiptAmount);
				ExecuteReceiptGoalSeek (workbook, num, receiptAmount);
				double num2 = ReadRequiredDouble (workbook, "お支払い履歴", "I" + num.ToString (CultureInfo.InvariantCulture), "残高", "AccountingPaymentHistory.OutputFutureBalance");
				while (num2 > receiptAmount) {
					num++;
					ApplyExpenseAmount (workbook, num, receiptAmount);
					ExecuteReceiptGoalSeek (workbook, num, receiptAmount);
					num2 = ReadRequiredDouble (workbook, "お支払い履歴", "I" + num.ToString (CultureInfo.InvariantCulture), "残高", "AccountingPaymentHistory.OutputFutureBalance");
				}
				if (num2 < 0.0) {
					CorrectRow (workbook, num);
					ClearRowsFrom (workbook, num + 1);
				} else {
					int num3 = num + 1;
					ApplyExpenseAmount (workbook, num3, receiptAmount);
					CorrectRow (workbook, num3);
					ClearRowsFrom (workbook, num3 + 1);
				}
				RefreshPrintArea (workbook);
				WritePaymentTotal (workbook);
				_logger.Info ("Payment history future balance output completed.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.OutputFutureBalance", exception);
			} finally {
				ProtectSheetSafely (workbook, "Payment history future balance reprotect failed.");
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState DeleteBlankReceiptRows (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "お支払い履歴");
				int editableStartRow = GetEditableStartRow (workbook);
				int startRow = CompactBlankDateRows (workbook, editableStartRow);
				if (editableStartRow == 13 && string.IsNullOrWhiteSpace (_accountingWorkbookService.ReadText (workbook, "お支払い履歴", "B13"))) {
					ClearRowsFrom (workbook, 13);
					_accountingWorkbookService.SetPrintAreaByBounds (workbook, "お支払い履歴", 13, 9);
					ClearPaymentTotal (workbook);
				} else {
					ClearRowsFrom (workbook, startRow);
					RefreshPrintArea (workbook);
					WritePaymentTotal (workbook);
				}
				_logger.Info ("Payment history blank receipt rows deleted.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.DeleteBlankReceiptRows", exception);
			} finally {
				ProtectSheetSafely (workbook, "Payment history delete reprotect failed.");
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState Reset (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, "お支払い履歴");
				SyncPaymentHistoryHeaderFromInvoice (workbook);
				_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "B12", "※");
				_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "A13", 1);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "発行日", "発行日");
				_accountingWorkbookService.ClearNamedRangeMergeAreaContents (workbook, "お支払い履歴", "請求額");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "お支払い履歴", "領収日");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "お支払い履歴", "分割金");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "お支払い履歴", "実費等総額");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "お支払い履歴", "源泉処理");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "お支払い履歴", "お預かり金額");
				_accountingWorkbookService.ClearNamedRangeContents (workbook, "お支払い履歴", "お預かり金額処理");
				_accountingWorkbookService.ClearRangeContents (workbook, "お支払い履歴", "B13:B72");
				ClearRowsFrom (workbook, 13);
				_accountingWorkbookService.SetPrintAreaByBounds (workbook, "お支払い履歴", 13, 9);
				ClearPaymentTotal (workbook);
				_logger.Info ("Payment history reset completed.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.Reset", exception);
			} finally {
				ProtectSheetSafely (workbook, "Payment history reset reprotect failed.");
			}
			return LoadFormState (workbook);
		}

		private bool TryValidateEntryRequest (AccountingPaymentHistoryEntryRequest request, out DateTime receiptDate, out double billingAmount, out double expenseAmount, out double depositAmount, out double receiptAmount)
		{
			receiptDate = DateTime.MinValue;
			billingAmount = ParseAmount (request.BillingAmountText);
			expenseAmount = ParseAmount (request.ExpenseAmountText);
			depositAmount = ParseAmount (request.DepositAmountText);
			receiptAmount = ParseAmount (request.ReceiptAmountText);
			if (billingAmount == 0.0) {
				ShowInformationMessage ("請求額が0円です。請求書を作成してください。");
				return false;
			}
			if (!TryParseDate (request.ReceiptDateText, out receiptDate)) {
				ShowInformationMessage ("領収日欄が日付になっていません。\r\n1973/02/10形式で入力し直してください。");
				return false;
			}
			if (string.IsNullOrWhiteSpace (request.ReceiptAmountText)) {
				ShowInformationMessage ("領収額が未入力です。");
				return false;
			}
			return true;
		}

		private bool TryValidateProjectionRequest (Workbook workbook, AccountingPaymentHistoryEntryRequest request, out double billingAmount, out double expenseAmount, out double receiptAmount)
		{
			billingAmount = ParseAmount (request.BillingAmountText);
			expenseAmount = ParseAmount (request.ExpenseAmountText);
			receiptAmount = ParseAmount (request.ReceiptAmountText);
			if (billingAmount == 0.0) {
				ShowInformationMessage ("請求額が0円です。請求書を作成してください。");
				return false;
			}
			if (string.IsNullOrWhiteSpace (request.ReceiptAmountText)) {
				ShowInformationMessage ("領収額が未入力です。");
				return false;
			}
			if (string.IsNullOrWhiteSpace (_accountingWorkbookService.ReadText (workbook, "お支払い履歴", "B13"))) {
				ShowInformationMessage ("1回目の領収日が未入力です。");
				return false;
			}
			int lastOccupiedHistoryRow = GetLastOccupiedHistoryRow (workbook, GetEditableStartRow (workbook));
			double num = ReadRequiredDouble (workbook, "お支払い履歴", "I" + lastOccupiedHistoryRow.ToString (CultureInfo.InvariantCulture), "残高", "AccountingPaymentHistory.OutputFutureBalance");
			double num2 = num / Math.Max (1.0, receiptAmount);
			double num3 = 60 - (lastOccupiedHistoryRow - 12) + 1;
			if (num2 > num3) {
				ShowInformationMessage ("分割回数が60回を超えてしまいます");
				return false;
			}
			return true;
		}

		private void WriteBaseValues (Workbook workbook, AccountingPaymentHistoryEntryRequest request, DateTime receiptDate, double billingAmount, double expenseAmount, double depositAmount, double receiptAmount)
		{
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "請求額", billingAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "実費等総額", expenseAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "源泉処理", request.WithholdingText ?? string.Empty);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "領収日", receiptDate);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "分割金", receiptAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "お預かり金額", depositAmount);
		}

		private void ApplyDepositRow (Workbook workbook, double depositAmount)
		{
			_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "A13", 0);
			_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "B13", "(充当済み)");
			_accountingWorkbookService.SetHorizontalAlignmentCenter (workbook, "お支払い履歴", "B13");
			double num = ReadRequiredDouble (workbook, "お支払い履歴", "J12", "実費残高", "AccountingPaymentHistory.AddHistoryEntry");
			double num2 = ((depositAmount >= num) ? num : depositAmount);
			_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "H13", num2);
			_accountingWorkbookService.ExecuteGoalSeek (workbook, "お支払い履歴", "C13", "D13", depositAmount);
			_accountingWorkbookService.RoundDownCell (workbook, "お支払い履歴", "D13", 0);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, "お支払い履歴", "お預かり金額処理", "済み");
		}

		private void WriteReceiptDateRow (Workbook workbook, int rowIndex, DateTime receiptDate)
		{
			_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "B" + rowIndex.ToString (CultureInfo.InvariantCulture), receiptDate);
		}

		private void ApplyExpenseAmount (Workbook workbook, int rowIndex, double targetAmount)
		{
			double num = ReadRequiredDouble (workbook, "お支払い履歴", "J" + (rowIndex - 1).ToString (CultureInfo.InvariantCulture), "実費残高", "AccountingPaymentHistory.ApplyExpenseAmount");
			double num2 = ((targetAmount >= num) ? num : targetAmount);
			_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "H" + rowIndex.ToString (CultureInfo.InvariantCulture), num2);
		}

		private void ExecuteReceiptGoalSeek (Workbook workbook, int rowIndex, double targetAmount)
		{
			string text = rowIndex.ToString (CultureInfo.InvariantCulture);
			_accountingWorkbookService.ExecuteGoalSeek (workbook, "お支払い履歴", "C" + text, "D" + text, targetAmount);
			_accountingWorkbookService.RoundDownCell (workbook, "お支払い履歴", "D" + text, 0);
		}

		private void CorrectRow (Workbook workbook, int rowIndex)
		{
			string text = rowIndex.ToString (CultureInfo.InvariantCulture);
			double num = ReadRequiredDouble (workbook, "お支払い履歴", "I" + (rowIndex - 1).ToString (CultureInfo.InvariantCulture), "残高", "AccountingPaymentHistory.CorrectRow");
			double num2 = ReadRequiredDouble (workbook, "お支払い履歴", "H" + text, "実費等", "AccountingPaymentHistory.CorrectRow");
			_accountingWorkbookService.ExecuteGoalSeek (workbook, "お支払い履歴", "F" + text, "D" + text, num - num2);
			_accountingWorkbookService.RoundDownCell (workbook, "お支払い履歴", "D" + text, 0);
		}

		private void ClearRowsFrom (Workbook workbook, int startRow)
		{
			string text = Math.Max (13, startRow).ToString (CultureInfo.InvariantCulture);
			string text2 = 73.ToString (CultureInfo.InvariantCulture);
			_accountingWorkbookService.ClearRangeContents (workbook, "お支払い履歴", "B" + text + ":B" + text2);
			_accountingWorkbookService.ClearRangeContents (workbook, "お支払い履歴", "D" + text + ":D" + text2);
			_accountingWorkbookService.ClearRangeContents (workbook, "お支払い履歴", "H" + text + ":H" + text2);
		}

		private int CompactBlankDateRows (Workbook workbook, int startRow)
		{
			Worksheet worksheet = null;
			try {
				worksheet = _accountingWorkbookService.GetWorksheet (workbook, "お支払い履歴");
				List<PaymentHistoryEditableRow> list = new List<PaymentHistoryEditableRow> ();
				for (int i = startRow; i <= 72; i++) {
					object value = ((_Worksheet)worksheet).get_Range ((object)("B" + i.ToString (CultureInfo.InvariantCulture)), Type.Missing).Value2;
					string value2 = Convert.ToString (value) ?? string.Empty;
					if (!string.IsNullOrWhiteSpace (value2)) {
						list.Add (new PaymentHistoryEditableRow (value, (dynamic)((_Worksheet)worksheet).get_Range ((object)("D" + i.ToString (CultureInfo.InvariantCulture)), Type.Missing).Value2, (dynamic)((_Worksheet)worksheet).get_Range ((object)("H" + i.ToString (CultureInfo.InvariantCulture)), Type.Missing).Value2, ResolveSortDate (value)));
					}
				}
				list.Sort ((PaymentHistoryEditableRow left, PaymentHistoryEditableRow right) => DateTime.Compare (left.SortDate, right.SortDate));
				ClearRowsFrom (workbook, startRow);
				int num = startRow;
				foreach (PaymentHistoryEditableRow item in list) {
					_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "B" + num.ToString (CultureInfo.InvariantCulture), item.DateValue);
					_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "D" + num.ToString (CultureInfo.InvariantCulture), item.TargetAmountValue);
					_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "H" + num.ToString (CultureInfo.InvariantCulture), item.ExpenseAmountValue);
					num++;
				}
				return num;
			} finally {
				ReleaseComObject (worksheet);
			}
		}

		private static DateTime ResolveSortDate (object value)
		{
			if (value is double d) {
				return DateTime.FromOADate (d);
			}
			if (value is DateTime result) {
				return result;
			}
			string s = Convert.ToString (value) ?? string.Empty;
			if (DateTime.TryParse (s, CultureInfo.CurrentCulture, DateTimeStyles.None, out var result2)) {
				return result2;
			}
			if (DateTime.TryParse (s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result3)) {
				return result3;
			}
			return DateTime.MaxValue;
		}

		private int GetNextAppendRow (Workbook workbook)
		{
			int lastOccupiedHistoryRow = GetLastOccupiedHistoryRow (workbook, GetEditableStartRow (workbook));
			return Math.Min (72, lastOccupiedHistoryRow + 1);
		}

		private int GetLastOccupiedHistoryRow (Workbook workbook, int startRow)
		{
			for (int num = 72; num >= startRow; num--) {
				string value = _accountingWorkbookService.ReadText (workbook, "お支払い履歴", "B" + num.ToString (CultureInfo.InvariantCulture));
				if (!string.IsNullOrWhiteSpace (value)) {
					return num;
				}
			}
			return startRow - 1;
		}

		private bool IsDepositProcessed (Workbook workbook)
		{
			string a = _accountingWorkbookService.ReadNamedRangeText (workbook, "お支払い履歴", "お預かり金額処理");
			return string.Equals (a, "済み", StringComparison.OrdinalIgnoreCase);
		}

		private void RefreshPrintArea (Workbook workbook)
		{
			int lastUsedRowInColumn = _accountingWorkbookService.GetLastUsedRowInColumn (workbook, "お支払い履歴", "D");
			if (lastUsedRowInColumn >= 1) {
				_accountingWorkbookService.SetPrintAreaByBounds (workbook, "お支払い履歴", lastUsedRowInColumn, 9);
			}
		}

		private void WritePaymentTotal (Workbook workbook)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				worksheet = _accountingWorkbookService.GetWorksheet (workbook, "お支払い履歴");
				range = ((_Worksheet)worksheet).get_Range ((object)("C13:C" + Math.Max (13, _accountingWorkbookService.GetLastUsedRowInColumn (workbook, "お支払い履歴", "C")).ToString (CultureInfo.InvariantCulture)), Type.Missing);
				object value = worksheet.Application.WorksheetFunction.Sum (range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				_accountingWorkbookService.WriteCellValue (workbook, "お支払い履歴", "I9", value);
			} finally {
				ReleaseComObject (range);
				ReleaseComObject (worksheet);
			}
		}

		private void ClearPaymentTotal (Workbook workbook)
		{
			_accountingWorkbookService.ClearRangeContents (workbook, "お支払い履歴", "I9");
		}

		private void SyncPaymentHistoryHeaderFromInvoice (Workbook workbook)
		{
			_accountingWorkbookService.CopyValueRange (workbook, "請求書", "A3:A4", "お支払い履歴", "A3:A4");
		}

		private static bool TryParseDate (string text, out DateTime date)
		{
			string s = (text ?? string.Empty).Trim ();
			return DateTime.TryParse (s, CultureInfo.CurrentCulture, DateTimeStyles.None, out date) || DateTime.TryParse (s, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
		}

		private static string FormatAmount (double amount)
		{
			return amount.ToString ("#,##0", CultureInfo.InvariantCulture);
		}

		private static double ParseAmount (string text)
		{
			string text2 = (text ?? string.Empty).Replace (",", string.Empty).Trim ();
			if (text2.Length == 0) {
				return 0.0;
			}
			if (double.TryParse (text2, NumberStyles.Any, CultureInfo.InvariantCulture, out var result)) {
				return result;
			}
			return 0.0;
		}

		private string ReadFormattedDateFromNamedRangeSafe (Workbook workbook, string sheetName, string rangeName)
		{
			try {
				string text = _accountingWorkbookService.ReadNamedRangeText (workbook, sheetName, rangeName);
				if (TryParseDate (text, out var date)) {
					return date.ToString ("yyyy/MM/dd", CultureInfo.InvariantCulture);
				}
				if (double.TryParse (text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result)) {
					return DateTime.FromOADate (result).ToString ("yyyy/MM/dd", CultureInfo.InvariantCulture);
				}
				return text;
			} catch {
				return string.Empty;
			}
		}

		private static void ReleaseComObject (object comObject)
		{
			CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (comObject);
		}

		private double ReadRequiredDouble (Workbook workbook, string sheetName, string address, string itemName, string procedureName)
		{
			return ReadNumericCellCore (workbook, sheetName, address, itemName, procedureName, allowBlankAsZero: false);
		}

		private double ReadDoubleAllowBlankAsZero (Workbook workbook, string sheetName, string address, string itemName, string procedureName)
		{
			return ReadNumericCellCore (workbook, sheetName, address, itemName, procedureName, allowBlankAsZero: true);
		}

		private double ReadLoadFormStateDoubleAllowBlankAsZero (Workbook workbook, string sheetName, string address, string itemName, string procedureName, List<string> warnings)
		{
			try {
				return ReadDoubleAllowBlankAsZero (workbook, sheetName, address, itemName, procedureName);
			} catch (InvalidOperationException exception) {
				if (warnings != null) {
					warnings.Add (exception.Message);
				}
				return 0.0;
			}
		}

		private double ReadNumericCellCore (Workbook workbook, string sheetName, string address, string itemName, string procedureName, bool allowBlankAsZero)
		{
			object cellValue = _accountingWorkbookService.ReadCellValue (workbook, sheetName, address);
			string displayText = _accountingWorkbookService.ReadDisplayText (workbook, sheetName, address);
			if (AccountingNumericCellReader.TryParseNumericCell (cellValue, displayText, out var value, out var isBlank)) {
				return value;
			}
			if (allowBlankAsZero && isBlank) {
				return 0.0;
			}
			InvalidOperationException ex = AccountingNumericCellReader.CreateReadFailureException (sheetName, address, itemName, procedureName, displayText, allowBlankAsZero);
			string text = Convert.ToString (cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
			_logger.Error ("Accounting numeric cell read failed. sheet=" + sheetName + ", address=" + address + ", item=" + itemName + ", procedure=" + procedureName + ", displayText=" + (string.IsNullOrWhiteSpace (displayText) ? "（空欄）" : displayText.Trim ()) + ", cellValue=" + text, ex);
			throw ex;
		}

		private bool ReadBooleanSafe (Workbook workbook, string sheetName, string address)
		{
			string a = _accountingWorkbookService.ReadText (workbook, sheetName, address);
			return string.Equals (a, "TRUE", StringComparison.OrdinalIgnoreCase) || string.Equals (a, "1", StringComparison.OrdinalIgnoreCase) || string.Equals (a, "-1", StringComparison.OrdinalIgnoreCase);
		}

		private void ProtectSheetSafely (Workbook workbook, string logMessage)
		{
			try {
				_accountingWorkbookService.ProtectSheetUiOnly (workbook, "お支払い履歴");
			} catch (Exception exception) {
				_logger.Error (logMessage, exception);
			}
		}

		private int GetEditableStartRow (Workbook workbook)
		{
			string a = _accountingWorkbookService.ReadText (workbook, "お支払い履歴", "B13");
			return string.Equals (a, "(充当済み)", StringComparison.OrdinalIgnoreCase) ? 14 : 13;
		}

		private void ShowLoadFormStateNumericReadWarning (string warningMessage)
		{
			if (string.IsNullOrWhiteSpace (warningMessage)) {
				return;
			}
			_logger.Warn ("AccountingPaymentHistory.LoadFormState numeric read warning. " + warningMessage.Replace (Environment.NewLine, " | "));
			MessageBox.Show ("数値読取に失敗した項目があります。該当項目は 0 として表示しています。" + Environment.NewLine + Environment.NewLine + warningMessage, "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}

		private static void ShowInformationMessage (string message)
		{
			MessageBox.Show (message ?? string.Empty, "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}
	}
}
