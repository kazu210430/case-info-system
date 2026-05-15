using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
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
			internal PaymentHistoryEditableRow (int originalRow, object dateValue, object taxBaseValue, object expenseChargeValue, DateTime sortDate, bool isDepositRow)
			{
				OriginalRow = originalRow;
				DateValue = dateValue;
				TaxBaseValue = taxBaseValue;
				ExpenseChargeValue = expenseChargeValue;
				SortDate = sortDate;
				IsDepositRow = isDepositRow;
			}

			internal int OriginalRow { get; private set; }

			internal object DateValue { get; private set; }

			internal object TaxBaseValue { get; private set; }

			internal object ExpenseChargeValue { get; private set; }

			internal DateTime SortDate { get; private set; }

			internal bool IsDepositRow { get; private set; }
		}

		private const string SheetName = "お支払い履歴";

		private const string InvoiceSheetName = "請求書";

		private const string PaymentHistoryIssueDateRangeName = "発行日";

		private const string PaymentHistoryBillingAmountRangeName = "請求額";

		private const string PaymentHistoryReceiptDateRangeName = "領収日";

		private const string PaymentHistoryAmountRangeName = "領収額";

		private const string PaymentHistoryExpenseAmountRangeName = "実費等総額";

		private const string PaymentHistoryWithholdingRangeName = "源泉処理";

		private const string PaymentHistoryDepositAmountRangeName = "お預かり金額";

		private const string PaymentTotalCellAddress = "I9";

		private const string MarkerCellAddress = "B12";

		private const string InvoiceBillingSubtotalCellAddress = "F23";

		private const string InvoiceExpenseCellAddress = "F25";

		private const string InvoiceWithholdingFlagCellAddress = "Y24";

		private const string InvoiceDepositAmountCellAddress = "F29";

		private const string IssueDateSourceCellAddress = "J1";

		private const string ResetMarkerText = "※";

		private const string IssueDatePlaceholderText = "発行日";

		private const string DateDisplayFormat = "yyyy/MM/dd";

		private const string TableRangeAddress = "A12:J73";

		private const double AmountTolerance = 0.01;
		private const string GoalSeekGenericUserMessage = "入力内容をご確認ください。";

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
			const string procedureName = "AccountingPaymentHistory.LoadFormState";
			List<string> warnings = new List<string> ();
			double billingSubtotal = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, InvoiceSheetName, InvoiceBillingSubtotalCellAddress, "請求額小計", procedureName, warnings);
			double expenseAmount = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, InvoiceSheetName, InvoiceExpenseCellAddress, "実費等総額", procedureName, warnings);
			double depositAmount = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, InvoiceSheetName, InvoiceDepositAmountCellAddress, "お預かり金額", procedureName, warnings);
			bool withholding = ReadBooleanSafe (workbook, InvoiceSheetName, InvoiceWithholdingFlagCellAddress);
			AccountingPaymentHistoryFormState state = new AccountingPaymentHistoryFormState {
				BillingAmountText = FormatAmount (billingSubtotal + expenseAmount),
				ExpenseAmountText = FormatAmount (expenseAmount),
				WithholdingText = (withholding ? "する" : "しない"),
				DepositAmountText = FormatAmount (depositAmount),
				ReceiptDateText = ReadFormattedDateFromNamedRangeSafe (workbook, SheetName, PaymentHistoryReceiptDateRangeName),
				ReceiptAmountText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, PaymentHistoryAmountRangeName),
				HasNumericReadError = warnings.Count > 0,
				NumericReadErrorMessage = string.Join (Environment.NewLine, warnings)
			};
			if (warnings.Count > 0) {
				ShowLoadFormStateNumericReadWarning (state.NumericReadErrorMessage);
			}
			return state;
		}

		internal AccountingPaymentHistoryFormState ApplyIssueDate (Workbook workbook)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, SheetName, IssueDateSourceCellAddress);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryIssueDateRangeName, value);
				_logger.Info ("Payment history issue date applied.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.ApplyIssueDate", exception);
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState ApplyToday (Workbook workbook)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, SheetName, IssueDateSourceCellAddress);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryReceiptDateRangeName, value);
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
			DateTime receiptDate;
			double billingAmount;
			double expenseAmount;
			double depositAmount;
			double receiptAmount;
			if (!TryValidateEntryRequest (request, out receiptDate, out billingAmount, out expenseAmount, out depositAmount, out receiptAmount)) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				EnsureWorkbookStructure (workbook);
				WriteBaseValues (workbook, request, receiptDate, billingAmount, expenseAmount, depositAmount, receiptAmount);
				CalculateWorkbook (workbook);

				int lastDataRow = SortPaymentHistoryRows (workbook);
				if (lastDataRow < AccountingPaymentHistoryPlanPolicy.FirstDataRow) {
					if (depositAmount != 0) {
						if (!TryWriteCalculatedRow (workbook, AccountingPaymentHistoryPlanPolicy.FirstDataRow, 0, AccountingPaymentHistoryPlanPolicy.DepositAppliedText, depositAmount, "お預かり金額")) {
							return LoadFormState (workbook);
						}
						if (!TryWriteCalculatedRow (workbook, AccountingPaymentHistoryPlanPolicy.FirstDataRow + 1, 1, receiptDate, receiptAmount, "領収額")) {
							return LoadFormState (workbook);
						}
					} else {
						if (!TryWriteCalculatedRow (workbook, AccountingPaymentHistoryPlanPolicy.FirstDataRow, 1, receiptDate, receiptAmount, "領収額")) {
							return LoadFormState (workbook);
						}
					}
				} else {
					int nextAppendRow = lastDataRow + 1;
					AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (nextAppendRow);
					if (!TryWriteCalculatedRow (workbook, nextAppendRow, ResolveNextRoundNumber (workbook, nextAppendRow), receiptDate, receiptAmount, "領収額")) {
						return LoadFormState (workbook);
					}
				}

				lastDataRow = SortPaymentHistoryRows (workbook);
				ClearRowsFrom (workbook, lastDataRow + 1);
				RefreshPrintArea (workbook, lastDataRow);
				RefreshPaymentTotal (workbook);
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
			DateTime receiptDate;
			double billingAmount;
			double expenseAmount;
			double depositAmount;
			double receiptAmount;
			if (!TryValidateProjectionRequest (request, out receiptDate, out billingAmount, out expenseAmount, out depositAmount, out receiptAmount)) {
				return LoadFormState (workbook);
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				EnsureWorkbookStructure (workbook);
				WriteBaseValues (workbook, request, receiptDate, billingAmount, expenseAmount, depositAmount, receiptAmount);
				CalculateWorkbook (workbook);

				int lastDataRow = SortPaymentHistoryRows (workbook);
				if (lastDataRow >= AccountingPaymentHistoryPlanPolicy.FirstDataRow) {
					double currentBalance = ReadRequiredDouble (workbook, SheetName, Address ("I", lastDataRow), "請求残高", "AccountingPaymentHistory.OutputFutureBalance");
					if (currentBalance <= AmountTolerance) {
						ClearRowsFrom (workbook, lastDataRow + 1);
						RefreshPrintArea (workbook, lastDataRow);
						RefreshPaymentTotal (workbook);
						return LoadFormState (workbook);
					}
				}

				int row = lastDataRow < AccountingPaymentHistoryPlanPolicy.FirstDataRow ? AccountingPaymentHistoryPlanPolicy.FirstDataRow : lastDataRow + 1;
				int lastWrittenRow = lastDataRow;
				while (true) {
					AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (row);
					if (!TryWriteCalculatedRow (workbook, row, ResolveNextRoundNumber (workbook, row), string.Empty, receiptAmount, "領収額")) {
						return LoadFormState (workbook);
					}
					lastWrittenRow = row;
					double balance = ReadRequiredDouble (workbook, SheetName, Address ("I", row), "請求残高", "AccountingPaymentHistory.OutputFutureBalance");
					if (balance > AmountTolerance) {
						row++;
						continue;
					}
					if (balance < -AmountTolerance) {
						if (!TryGoalSeekAndVerify (workbook, row, "I", "D", 0, "最終回の請求残高")) {
							return LoadFormState (workbook);
						}
					}
					break;
				}

				ClearRowsFrom (workbook, lastWrittenRow + 1);
				RefreshPrintArea (workbook, lastWrittenRow);
				RefreshPaymentTotal (workbook);
				_logger.Info ("Payment history future balance output completed.");
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.OutputFutureBalance", exception);
			} finally {
				ProtectSheetSafely (workbook, "Payment history future balance reprotect failed.");
			}
			return LoadFormState (workbook);
		}

		internal AccountingPaymentHistoryFormState DeleteSelectedRows (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				EnsureWorkbookStructure (workbook);
				int selectedRowCount = ClearReceiptDatesForSelectedTableRows (workbook);
				int lastDataRow = SortPaymentHistoryRows (workbook);
				ClearRowsFrom (workbook, lastDataRow + 1);
				RefreshPrintArea (workbook, lastDataRow);
				RefreshPaymentTotal (workbook);
				_logger.Info ("Payment history selected rows deleted. selectedRows=" + selectedRowCount.ToString (CultureInfo.InvariantCulture));
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingPaymentHistory.DeleteSelectedRows", exception);
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
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				EnsureWorkbookStructure (workbook);
				SyncPaymentHistoryHeaderFromInvoice (workbook);
				_accountingWorkbookService.WriteCellValue (workbook, SheetName, MarkerCellAddress, ResetMarkerText);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryIssueDateRangeName, IssueDatePlaceholderText);
				_accountingWorkbookService.ClearNamedRangeMergeAreaContents (workbook, SheetName, PaymentHistoryBillingAmountRangeName);
				_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, PaymentHistoryReceiptDateRangeName);
				_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, PaymentHistoryAmountRangeName);
				_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, PaymentHistoryExpenseAmountRangeName);
				_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, PaymentHistoryWithholdingRangeName);
				_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, PaymentHistoryDepositAmountRangeName);
				ClearRowsFrom (workbook, AccountingPaymentHistoryPlanPolicy.FirstDataRow);
				RefreshPrintArea (workbook, AccountingPaymentHistoryPlanPolicy.StartValueRow);
				RefreshPaymentTotal (workbook);
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
			return TryValidateCommonRequest (request, requireReceiptDate: true, receiptDate: out receiptDate, billingAmount: out billingAmount, expenseAmount: out expenseAmount, depositAmount: out depositAmount, receiptAmount: out receiptAmount);
		}

		private bool TryValidateProjectionRequest (AccountingPaymentHistoryEntryRequest request, out DateTime receiptDate, out double billingAmount, out double expenseAmount, out double depositAmount, out double receiptAmount)
		{
			return TryValidateCommonRequest (request, requireReceiptDate: true, receiptDate: out receiptDate, billingAmount: out billingAmount, expenseAmount: out expenseAmount, depositAmount: out depositAmount, receiptAmount: out receiptAmount);
		}

		private bool TryValidateCommonRequest (AccountingPaymentHistoryEntryRequest request, bool requireReceiptDate, out DateTime receiptDate, out double billingAmount, out double expenseAmount, out double depositAmount, out double receiptAmount)
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
			if (expenseAmount < 0.0) {
				ShowInformationMessage ("実費等総額にマイナス金額は指定できません。");
				return false;
			}
			if (depositAmount < 0.0) {
				ShowInformationMessage ("お預かり金額にマイナス金額は指定できません。");
				return false;
			}
			if (requireReceiptDate && !TryParseDate (request.ReceiptDateText, out receiptDate)) {
				ShowInformationMessage ("領収日欄が日付になっていません。\r\n1973/02/10形式で入力し直してください。");
				return false;
			}
			if (string.IsNullOrWhiteSpace (request.ReceiptAmountText)) {
				ShowInformationMessage ("領収額が未入力です。");
				return false;
			}
			if (receiptAmount <= 0.0) {
				ShowInformationMessage ("領収額は1円以上で入力してください。");
				return false;
			}
			return true;
		}

		private void WriteBaseValues (Workbook workbook, AccountingPaymentHistoryEntryRequest request, DateTime receiptDate, double billingAmount, double expenseAmount, double depositAmount, double receiptAmount)
		{
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryBillingAmountRangeName, billingAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryExpenseAmountRangeName, expenseAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryWithholdingRangeName, request.WithholdingText ?? string.Empty);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryReceiptDateRangeName, receiptDate);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryAmountRangeName, receiptAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, PaymentHistoryDepositAmountRangeName, depositAmount);
		}

		private bool TryWriteCalculatedRow (Workbook workbook, int rowIndex, int roundNumber, object dateValue, double targetAmount, string targetName)
		{
			AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (rowIndex);
			int previousRow = AccountingPaymentHistoryPlanPolicy.GetPreviousBalanceRowForDataRow (rowIndex);
			double previousExpenseBalance = ReadRequiredDouble (workbook, SheetName, Address ("J", previousRow), "実費残高", "AccountingPaymentHistory.WriteCalculatedRow");
			double expenseCharge = AccountingPaymentHistoryPlanPolicy.ResolveExpenseCharge (targetAmount, previousExpenseBalance);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("A", rowIndex), roundNumber);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("B", rowIndex), dateValue);
			if (dateValue is string && AccountingPaymentHistoryPlanPolicy.IsDepositMarker ((string)dateValue)) {
				_accountingWorkbookService.SetHorizontalAlignmentCenter (workbook, SheetName, Address ("B", rowIndex));
			}
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("D", rowIndex), 0);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("H", rowIndex), expenseCharge);
			CalculateWorkbook (workbook);
			return TryGoalSeekAndVerify (workbook, rowIndex, "C", "D", targetAmount, targetName);
		}

		private bool TryGoalSeekAndVerify (Workbook workbook, int rowIndex, string formulaColumn, string changingColumn, double targetAmount, string targetName)
		{
			string formulaCell = Address (formulaColumn, rowIndex);
			string changingCell = Address (changingColumn, rowIndex);
			AccountingGoalSeekExecutionResult goalSeekResult;
			try {
				goalSeekResult = _accountingWorkbookService.ExecuteGoalSeekAndReadResult (workbook, SheetName, formulaCell, changingCell, targetAmount);
			} catch (Exception exception) {
				throw CreateGoalSeekFailureException (workbook, rowIndex, formulaCell, changingCell, targetAmount, null, null, null, targetName, GoalSeekGenericUserMessage, "GoalSeek 実行中に例外が発生しました。", exception);
			}
			if (!goalSeekResult.Succeeded) {
				double resultCurrent;
				double? current = TryGetGoalSeekCurrentFromResult (goalSeekResult, out resultCurrent) ? resultCurrent : (double?)null;
				if (current.HasValue && AccountingGoalSeekResidualPolicy.IsWithinAllowedResidual (current.Value, targetAmount)) {
					_logger.Warn (CreateGoalSeekDiagnostic (workbook, rowIndex, formulaCell, changingCell, targetAmount, current, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, "GoalSeek が false を返しましたが残差は 1 円未満です。", null));
				} else if (current.HasValue) {
					ShowGoalSeekResidualNotice (workbook, rowIndex, formulaCell, changingCell, targetAmount, current.Value, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, "GoalSeek が false を返し、1 円以上の残差が残りました。");
					return false;
				} else {
					throw CreateGoalSeekFailureException (workbook, rowIndex, formulaCell, changingCell, targetAmount, current, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, GoalSeekGenericUserMessage, "GoalSeek が false を返し、目標セル値を数値として取得できません。", null);
				}
			}
			CalculateWorkbook (workbook);

			double actual;
			try {
				actual = ReadRequiredDouble (workbook, SheetName, formulaCell, targetName, "AccountingPaymentHistory.GoalSeekAndVerify");
			} catch (Exception exception) {
				throw CreateGoalSeekFailureException (workbook, rowIndex, formulaCell, changingCell, targetAmount, null, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, GoalSeekGenericUserMessage, "GoalSeek 後の目標セル値を数値として読み取れません。", exception);
			}

			if (AccountingGoalSeekResidualPolicy.ShouldShowResidualNotice (actual, targetAmount)) {
				ShowGoalSeekResidualNotice (workbook, rowIndex, formulaCell, changingCell, targetAmount, actual, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, "GoalSeek 後に 1 円以上の残差が残りました。");
				return false;
			}

			double taxBase;
			try {
				taxBase = ReadRequiredDouble (workbook, SheetName, changingCell, "10％対象計", "AccountingPaymentHistory.GoalSeekAndVerify");
			} catch (Exception exception) {
				throw CreateGoalSeekFailureException (workbook, rowIndex, formulaCell, changingCell, targetAmount, actual, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, GoalSeekGenericUserMessage, "GoalSeek 後の変更セル値を数値として読み取れません。", exception);
			}
			if (taxBase < -AmountTolerance) {
				throw CreateGoalSeekFailureException (workbook, rowIndex, formulaCell, changingCell, targetAmount, actual, goalSeekResult.CurrentValue, goalSeekResult.Succeeded, targetName, "10％対象計がマイナスになっているのでご確認ください。", "GoalSeek 後の10％対象計がマイナスになりました。taxBase=" + taxBase.ToString (CultureInfo.InvariantCulture), null);
			}

			return true;
		}

		private void ShowGoalSeekResidualNotice (Workbook workbook, int rowIndex, string formulaCell, string changingCell, double targetAmount, double current, object rawCurrent, bool? goalSeekSucceeded, string targetName, string reason)
		{
			_logger.Warn (CreateGoalSeekDiagnostic (workbook, rowIndex, formulaCell, changingCell, targetAmount, current, rawCurrent, goalSeekSucceeded, targetName, reason, null));
			UserErrorService.ShowOkNotification (AccountingGoalSeekResidualPolicy.CreateResidualNoticeUserMessage (current, targetAmount), "案件情報System", MessageBoxIcon.Warning);
		}

		private int SortPaymentHistoryRows (Workbook workbook)
		{
			List<PaymentHistoryEditableRow> rows = ReadEditableRows (workbook);
			PaymentHistoryEditableRow depositRow = null;
			List<PaymentHistoryEditableRow> receiptRows = new List<PaymentHistoryEditableRow> ();
			foreach (PaymentHistoryEditableRow row in rows) {
				if (row.IsDepositRow) {
					if (depositRow != null) {
						throw new InvalidOperationException ("お支払い履歴に0回目の充当行が複数あります。B列の「（充当済み）」行を確認してください。");
					}
					depositRow = row;
				} else {
					receiptRows.Add (row);
				}
			}

			receiptRows.Sort (CompareEditableRows);
			ClearRowsFrom (workbook, AccountingPaymentHistoryPlanPolicy.FirstDataRow);
			int targetRow = AccountingPaymentHistoryPlanPolicy.FirstDataRow;
			int roundNumber = 1;
			if (depositRow != null) {
				WriteEditableRowValues (workbook, targetRow, 0, AccountingPaymentHistoryPlanPolicy.DepositAppliedText, depositRow.TaxBaseValue, depositRow.ExpenseChargeValue);
				targetRow++;
			}
			foreach (PaymentHistoryEditableRow row in receiptRows) {
				WriteEditableRowValues (workbook, targetRow, roundNumber, row.DateValue, row.TaxBaseValue, row.ExpenseChargeValue);
				targetRow++;
				roundNumber++;
			}
			CalculateWorkbook (workbook);
			return targetRow - 1;
		}

		private List<PaymentHistoryEditableRow> ReadEditableRows (Workbook workbook)
		{
			List<PaymentHistoryEditableRow> rows = new List<PaymentHistoryEditableRow> ();
			for (int row = AccountingPaymentHistoryPlanPolicy.FirstDataRow; row <= AccountingPaymentHistoryPlanPolicy.LastDataRow; row++) {
				object dateValue = _accountingWorkbookService.ReadCellValue (workbook, SheetName, Address ("B", row));
				string displayText = _accountingWorkbookService.ReadDisplayText (workbook, SheetName, Address ("B", row));
				if (IsBlankCellValue (dateValue) && string.IsNullOrWhiteSpace (displayText)) {
					continue;
				}
				bool isDepositRow = AccountingPaymentHistoryPlanPolicy.IsDepositMarker (displayText) || AccountingPaymentHistoryPlanPolicy.IsDepositMarker (Convert.ToString (dateValue, CultureInfo.InvariantCulture));
				DateTime sortDate = isDepositRow ? DateTime.MinValue : ResolveSortDate (dateValue, displayText, row);
				rows.Add (new PaymentHistoryEditableRow (
					row,
					isDepositRow ? (object)AccountingPaymentHistoryPlanPolicy.DepositAppliedText : dateValue,
					_accountingWorkbookService.ReadCellValue (workbook, SheetName, Address ("D", row)),
					_accountingWorkbookService.ReadCellValue (workbook, SheetName, Address ("H", row)),
					sortDate,
					isDepositRow));
			}
			return rows;
		}

		private static int CompareEditableRows (PaymentHistoryEditableRow left, PaymentHistoryEditableRow right)
		{
			int result = DateTime.Compare (left.SortDate, right.SortDate);
			if (result != 0) {
				return result;
			}
			return left.OriginalRow.CompareTo (right.OriginalRow);
		}

		private void WriteEditableRowValues (Workbook workbook, int rowIndex, int roundNumber, object dateValue, object taxBaseValue, object expenseChargeValue)
		{
			AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (rowIndex);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("A", rowIndex), roundNumber);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("B", rowIndex), dateValue);
			if (dateValue is string && AccountingPaymentHistoryPlanPolicy.IsDepositMarker ((string)dateValue)) {
				_accountingWorkbookService.SetHorizontalAlignmentCenter (workbook, SheetName, Address ("B", rowIndex));
			}
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("D", rowIndex), taxBaseValue);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, Address ("H", rowIndex), expenseChargeValue);
		}

		private int ResolveNextRoundNumber (Workbook workbook, int rowIndex)
		{
			if (rowIndex <= AccountingPaymentHistoryPlanPolicy.FirstDataRow) {
				return 1;
			}
			double previousRound = ReadRequiredDouble (workbook, SheetName, Address ("A", rowIndex - 1), "回数", "AccountingPaymentHistory.ResolveNextRoundNumber");
			return Math.Max (0, Convert.ToInt32 (Math.Round (previousRound, 0))) + 1;
		}

		private void ClearRowsFrom (Workbook workbook, int startRow)
		{
			int row = Math.Max (AccountingPaymentHistoryPlanPolicy.FirstDataRow, startRow);
			if (row > AccountingPaymentHistoryPlanPolicy.LastDataRow) {
				return;
			}
			string start = row.ToString (CultureInfo.InvariantCulture);
			string end = AccountingPaymentHistoryPlanPolicy.LastDataRow.ToString (CultureInfo.InvariantCulture);
			_accountingWorkbookService.ClearRangeContents (workbook, SheetName, "B" + start + ":B" + end);
			_accountingWorkbookService.ClearRangeContents (workbook, SheetName, "D" + start + ":D" + end);
			_accountingWorkbookService.ClearRangeContents (workbook, SheetName, "H" + start + ":H" + end);
		}

		private int ClearReceiptDatesForSelectedTableRows (Workbook workbook)
		{
			List<int> selectedRows = ResolveSelectedPaymentHistoryRows (workbook);
			if (selectedRows.Count == 0) {
				return 0;
			}

			foreach (int row in selectedRows) {
				_accountingWorkbookService.ClearRangeContents (workbook, SheetName, Address ("B", row));
			}

			_logger.Info ("Payment history selected receipt dates cleared. rows=" + string.Join (",", selectedRows));
			return selectedRows.Count;
		}

		private List<int> ResolveSelectedPaymentHistoryRows (Workbook workbook)
		{
			List<int> selectedRows = new List<int> ();
			Range selection = null;
			Worksheet selectedWorksheet = null;
			Worksheet worksheet = null;
			Range dataRange = null;
			Range intersection = null;
			Areas areas = null;

			try {
				selection = workbook != null && workbook.Application != null ? workbook.Application.Selection as Range : null;
				if (selection == null) {
					return selectedRows;
				}

				selectedWorksheet = selection.Worksheet as Worksheet;
				if (selectedWorksheet == null || !string.Equals (selectedWorksheet.Name, SheetName, StringComparison.Ordinal)) {
					return selectedRows;
				}

				worksheet = _accountingWorkbookService.GetWorksheet (workbook, SheetName);
				dataRange = worksheet.Range["A" + AccountingPaymentHistoryPlanPolicy.FirstDataRow.ToString (CultureInfo.InvariantCulture) + ":J" + AccountingPaymentHistoryPlanPolicy.LastDataRow.ToString (CultureInfo.InvariantCulture)];
				intersection = worksheet.Application.Intersect (
					selection,
					dataRange,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing);

				if (intersection == null) {
					return selectedRows;
				}

				areas = intersection.Areas;
				for (int areaIndex = 1; areaIndex <= areas.Count; areaIndex++) {
					Range area = null;
					Range areaRows = null;
					try {
						area = areas[areaIndex];
						areaRows = area.Rows;
						int firstRow = area.Row;
						int lastRow = Math.Min (AccountingPaymentHistoryPlanPolicy.LastDataRow, firstRow + areaRows.Count - 1);

						for (int selectedRow = Math.Max (AccountingPaymentHistoryPlanPolicy.FirstDataRow, firstRow); selectedRow <= lastRow; selectedRow++) {
							if (!selectedRows.Contains (selectedRow)) {
								selectedRows.Add (selectedRow);
							}
						}
					} finally {
						ComObjectReleaseService.Release (areaRows);
						ComObjectReleaseService.Release (area);
					}
				}

				selectedRows.Sort ();
				return selectedRows;
			} finally {
				ComObjectReleaseService.Release (areas);
				ComObjectReleaseService.Release (intersection);
				ComObjectReleaseService.Release (dataRange);
				ComObjectReleaseService.Release (worksheet);
				ComObjectReleaseService.Release (selectedWorksheet);
			}
		}

		private void RefreshPrintArea (Workbook workbook, int lastDataRow)
		{
			int printLastRow = Math.Max (AccountingPaymentHistoryPlanPolicy.StartValueRow, Math.Min (lastDataRow, AccountingPaymentHistoryPlanPolicy.LastDataRow));
			_accountingWorkbookService.SetPrintAreaByBounds (workbook, SheetName, printLastRow, AccountingPaymentHistoryPlanPolicy.PrintLastColumn);
		}

		private void RefreshPaymentTotal (Workbook workbook)
		{
			EnsureFormulaEquals (workbook, PaymentTotalCellAddress, "=SUM(C14:C73)", "既払い額計");
			CalculateWorkbook (workbook);
		}

		private void EnsureWorkbookStructure (Workbook workbook)
		{
			EnsureNamedRangeAddress (workbook, PaymentHistoryBillingAmountRangeName, "I7");
			EnsureNamedRangeAddress (workbook, PaymentHistoryReceiptDateRangeName, "K5");
			EnsureNamedRangeAddress (workbook, PaymentHistoryAmountRangeName, "K3");
			EnsureNamedRangeAddress (workbook, PaymentHistoryExpenseAmountRangeName, "K7");
			EnsureNamedRangeAddress (workbook, PaymentHistoryWithholdingRangeName, "K9");
			EnsureNamedRangeAddress (workbook, PaymentHistoryDepositAmountRangeName, "K11");
			EnsureNamedRangeAddress (workbook, "既払い額計", "I9");

			if (!_accountingWorkbookService.HasListObjectRange (workbook, SheetName, TableRangeAddress)) {
				throw new InvalidOperationException ("お支払い履歴のテーブル範囲 A12:J73 が見つかりません。ブック側を先に確認してください。");
			}

			EnsureFormulaEquals (workbook, "I13", "=" + PaymentHistoryBillingAmountRangeName, "13行目 請求額開始値");
			EnsureFormulaEquals (workbook, "J13", "=" + PaymentHistoryExpenseAmountRangeName, "13行目 実費等総額開始値");
			EnsureFormulaEquals (workbook, PaymentTotalCellAddress, "=SUM(C14:C73)", "既払い額計");
			for (int row = AccountingPaymentHistoryPlanPolicy.FirstDataRow; row <= AccountingPaymentHistoryPlanPolicy.LastDataRow; row++) {
				EnsureFormulaPresent (workbook, Address ("C", row), "領収額");
				EnsureFormulaPresent (workbook, Address ("E", row), "10％消費税");
				EnsureFormulaPresent (workbook, Address ("F", row), "小計");
				EnsureFormulaPresent (workbook, Address ("G", row), "源泉徴収税");
				EnsureFormulaPresent (workbook, Address ("I", row), "請求残高");
				EnsureFormulaPresent (workbook, Address ("J", row), "実費残高");
			}
		}

		private void EnsureNamedRangeAddress (Workbook workbook, string rangeName, string expectedAddress)
		{
			string actualAddress = NormalizeAddress (_accountingWorkbookService.GetNamedRangeAddress (workbook, SheetName, rangeName));
			if (!string.Equals (actualAddress, expectedAddress, StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("お支払い履歴の名付セル [" + rangeName + "] の参照先が想定と違います。想定: " + expectedAddress + ", 実際: " + actualAddress);
			}
		}

		private void EnsureFormulaPresent (Workbook workbook, string address, string itemName)
		{
			string formula = ReadFormulaText (workbook, address);
			if (string.IsNullOrWhiteSpace (formula) || !formula.TrimStart ().StartsWith ("=", StringComparison.Ordinal)) {
				throw new InvalidOperationException ("お支払い履歴!" + address + " の " + itemName + " の式が見つかりません。ブック側の式を先に確認してください。");
			}
		}

		private void EnsureFormulaEquals (Workbook workbook, string address, string expectedFormula, string itemName)
		{
			string actualFormula = NormalizeFormula (ReadFormulaText (workbook, address));
			string expected = NormalizeFormula (expectedFormula);
			if (!string.Equals (actualFormula, expected, StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("お支払い履歴!" + address + " の " + itemName + " の式が想定と違います。想定: " + expectedFormula + ", 実際: " + ReadFormulaText (workbook, address));
			}
		}

		private string ReadFormulaText (Workbook workbook, string address)
		{
			object formula = _accountingWorkbookService.ReadRangeFormulaSnapshot (workbook, SheetName, address);
			return Convert.ToString (formula, CultureInfo.InvariantCulture) ?? string.Empty;
		}

		private UserFacingException CreateGoalSeekFailureException (Workbook workbook, int rowIndex, string formulaCell, string changingCell, double targetAmount, double? current, object rawCurrent, bool? goalSeekSucceeded, string targetName, string userMessage, string reason, Exception exception)
		{
			string diagnostic = CreateGoalSeekDiagnostic (workbook, rowIndex, formulaCell, changingCell, targetAmount, current, rawCurrent, goalSeekSucceeded, targetName, reason, exception);
			return exception == null ? new UserFacingException (userMessage, diagnostic) : new UserFacingException (userMessage, diagnostic, exception);
		}

		private string CreateGoalSeekDiagnostic (Workbook workbook, int rowIndex, string formulaCell, string changingCell, double targetAmount, double? current, object rawCurrent, bool? goalSeekSucceeded, string targetName, string reason, Exception exception)
		{
			double? residual = current.HasValue ? AccountingGoalSeekResidualPolicy.GetResidual (current.Value, targetAmount) : (double?)null;
			double? residualAbs = current.HasValue ? AccountingGoalSeekResidualPolicy.GetResidualAbs (current.Value, targetAmount) : (double?)null;
			StringBuilder builder = new StringBuilder ();
			builder.AppendLine ("AccountingPaymentHistory.GoalSeek diagnostic. reason=" + (reason ?? string.Empty));
			builder.AppendLine ("procedure=AccountingPaymentHistory.GoalSeekAndVerify");
			builder.AppendLine ("sheet=" + SheetName);
			builder.AppendLine ("対象行: " + rowIndex.ToString (CultureInfo.InvariantCulture));
			builder.AppendLine ("targetName=" + (targetName ?? string.Empty));
			builder.AppendLine ("formulaCell=" + (formulaCell ?? string.Empty) + ", changingCell=" + (changingCell ?? string.Empty));
			builder.AppendLine ("target=" + targetAmount.ToString (CultureInfo.InvariantCulture) + ", current=" + FormatNullableDouble (current) + ", rawCurrent=" + FormatObjectForLog (rawCurrent));
			builder.AppendLine ("residual=" + FormatNullableDouble (residual) + ", residualAbs=" + FormatNullableDouble (residualAbs) + ", residualYen=" + (residualAbs.HasValue ? AccountingGoalSeekResidualPolicy.FormatResidualYen (residualAbs.Value) : "(unavailable)"));
			builder.AppendLine ("goalSeekResult=" + FormatNullableBool (goalSeekSucceeded));
			if (exception != null) {
				builder.AppendLine ("exceptionType=" + exception.GetType ().FullName + ", exceptionMessage=" + exception.Message);
			}
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "C", "領収額"));
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "D", "10％対象計"));
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "F", "小計"));
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "G", "源泉徴収税"));
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "H", "実費等への充当額"));
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "I", "請求残高"));
			builder.AppendLine (ReadDiagnosticCell (workbook, rowIndex, "J", "実費残高"));
			return builder.ToString ();
		}

		private static bool TryGetGoalSeekCurrentFromResult (AccountingGoalSeekExecutionResult goalSeekResult, out double current)
		{
			current = 0;
			if (goalSeekResult == null || !(goalSeekResult.CurrentValue is double)) {
				return false;
			}
			current = (double)goalSeekResult.CurrentValue;
			return true;
		}

		private static string FormatNullableDouble (double? value)
		{
			return value.HasValue ? value.Value.ToString (CultureInfo.InvariantCulture) : "(unavailable)";
		}

		private static string FormatNullableBool (bool? value)
		{
			return value.HasValue ? value.Value.ToString () : "(unknown)";
		}

		private static string FormatObjectForLog (object value)
		{
			return Convert.ToString (value, CultureInfo.InvariantCulture) ?? string.Empty;
		}

		private string ReadDiagnosticCell (Workbook workbook, int rowIndex, string column, string itemName)
		{
			string address = Address (column, rowIndex);
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, SheetName, address);
				string display = _accountingWorkbookService.ReadDisplayText (workbook, SheetName, address);
				string formula = ReadFormulaText (workbook, address);
				string formulaState = string.IsNullOrWhiteSpace (formula) || !formula.TrimStart ().StartsWith ("=", StringComparison.Ordinal) ? "式なし" : "式あり";
				return address + " " + itemName + ": value=" + (Convert.ToString (value, CultureInfo.InvariantCulture) ?? string.Empty) + ", display=" + display + ", " + formulaState;
			} catch (Exception exception) {
				return address + " " + itemName + ": 診断取得失敗 " + exception.Message;
			}
		}

		private static DateTime ResolveSortDate (object value, string displayText, int rowIndex)
		{
			if (value is double) {
				return DateTime.FromOADate ((double)value);
			}
			if (value is DateTime) {
				return (DateTime)value;
			}
			string text = (displayText ?? string.Empty).Trim ();
			DateTime result;
			if (DateTime.TryParse (text, CultureInfo.CurrentCulture, DateTimeStyles.None, out result)) {
				return result;
			}
			if (DateTime.TryParse (text, CultureInfo.InvariantCulture, DateTimeStyles.None, out result)) {
				return result;
			}
			string valueText = Convert.ToString (value, CultureInfo.InvariantCulture) ?? string.Empty;
			if (DateTime.TryParse (valueText, CultureInfo.CurrentCulture, DateTimeStyles.None, out result)) {
				return result;
			}
			if (DateTime.TryParse (valueText, CultureInfo.InvariantCulture, DateTimeStyles.None, out result)) {
				return result;
			}
			throw new InvalidOperationException ("お支払い履歴!B" + rowIndex.ToString (CultureInfo.InvariantCulture) + " が日付として読めません。B列には日付または「（充当済み）」を入力してください。");
		}

		private static bool IsBlankCellValue (object value)
		{
			if (value == null) {
				return true;
			}
			string text = Convert.ToString (value, CultureInfo.InvariantCulture);
			return string.IsNullOrWhiteSpace (text);
		}

		private static void CalculateWorkbook (Workbook workbook)
		{
			if (workbook != null && workbook.Application != null) {
				workbook.Application.Calculate ();
			}
		}

		private void SyncPaymentHistoryHeaderFromInvoice (Workbook workbook)
		{
			_accountingWorkbookService.CopyValueRange (workbook, InvoiceSheetName, "A3:A4", SheetName, "A3:A4");
		}

		private static bool TryParseDate (string text, out DateTime date)
		{
			string value = (text ?? string.Empty).Trim ();
			return DateTime.TryParse (value, CultureInfo.CurrentCulture, DateTimeStyles.None, out date) || DateTime.TryParse (value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
		}

		private static string FormatAmount (double amount)
		{
			return amount.ToString ("#,##0", CultureInfo.InvariantCulture);
		}

		private static double ParseAmount (string text)
		{
			string value = (text ?? string.Empty).Replace (",", string.Empty).Trim ();
			if (value.Length == 0) {
				return 0.0;
			}
			double result;
			if (double.TryParse (value, NumberStyles.Any, CultureInfo.InvariantCulture, out result)) {
				return result;
			}
			if (double.TryParse (value, NumberStyles.Any, CultureInfo.CurrentCulture, out result)) {
				return result;
			}
			return 0.0;
		}

		private string ReadFormattedDateFromNamedRangeSafe (Workbook workbook, string sheetName, string rangeName)
		{
			try {
				string text = _accountingWorkbookService.ReadNamedRangeText (workbook, sheetName, rangeName);
				DateTime date;
				if (TryParseDate (text, out date)) {
					return date.ToString (DateDisplayFormat, CultureInfo.InvariantCulture);
				}
				double result;
				if (double.TryParse (text, NumberStyles.Any, CultureInfo.InvariantCulture, out result)) {
					return DateTime.FromOADate (result).ToString (DateDisplayFormat, CultureInfo.InvariantCulture);
				}
				return text;
			} catch {
				return string.Empty;
			}
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
			double value;
			bool isBlank;
			if (AccountingNumericCellReader.TryParseNumericCell (cellValue, displayText, out value, out isBlank)) {
				return value;
			}
			if (allowBlankAsZero && isBlank) {
				return 0.0;
			}
			InvalidOperationException exception = AccountingNumericCellReader.CreateReadFailureException (sheetName, address, itemName, procedureName, displayText, allowBlankAsZero);
			string text = Convert.ToString (cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
			_logger.Error ("Accounting numeric cell read failed. sheet=" + sheetName + ", address=" + address + ", item=" + itemName + ", procedure=" + procedureName + ", displayText=" + (string.IsNullOrWhiteSpace (displayText) ? "（空欄）" : displayText.Trim ()) + ", cellValue=" + text, exception);
			throw exception;
		}

		private bool ReadBooleanSafe (Workbook workbook, string sheetName, string address)
		{
			string value = _accountingWorkbookService.ReadText (workbook, sheetName, address);
			return string.Equals (value, "TRUE", StringComparison.OrdinalIgnoreCase) || string.Equals (value, "1", StringComparison.OrdinalIgnoreCase) || string.Equals (value, "-1", StringComparison.OrdinalIgnoreCase);
		}

		private void ProtectSheetSafely (Workbook workbook, string logMessage)
		{
			try {
				_accountingWorkbookService.ProtectSheetUiOnly (workbook, SheetName);
			} catch (Exception exception) {
				_logger.Error (logMessage, exception);
			}
		}

		private void ShowLoadFormStateNumericReadWarning (string warningMessage)
		{
			if (string.IsNullOrWhiteSpace (warningMessage)) {
				return;
			}
			_logger.Warn ("AccountingPaymentHistory.LoadFormState numeric read warning. " + warningMessage.Replace (Environment.NewLine, " | "));
			UserErrorService.ShowOkNotification ("数値読取に失敗した項目があります。該当項目は 0 として表示しています。" + Environment.NewLine + Environment.NewLine + warningMessage, "案件情報System", MessageBoxIcon.Warning);
		}

		private static string Address (string column, int row)
		{
			return column + row.ToString (CultureInfo.InvariantCulture);
		}

		private static string NormalizeAddress (string address)
		{
			return (address ?? string.Empty).Replace ("$", string.Empty).Trim ();
		}

		private static string NormalizeFormula (string formula)
		{
			return (formula ?? string.Empty).Replace ("$", string.Empty).Replace (" ", string.Empty).Trim ();
		}

		private static void ShowInformationMessage (string message)
		{
			UserErrorService.ShowOkNotification (message ?? string.Empty, "案件情報System", MessageBoxIcon.Asterisk);
		}
	}
}
