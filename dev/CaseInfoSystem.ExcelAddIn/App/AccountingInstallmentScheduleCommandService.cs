using System;
using System.Collections.Generic;
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
		private const string InvoiceSheetName = "請求書";
		private const string TableRangeAddress = "A12:J73";

		private const string InstallmentIssueDateRangeName = "発行日";
		private const string InstallmentBillingAmountRangeName = "請求額";
		private const string InstallmentExpenseAmountRangeName = "実費等総額";
		private const string InstallmentWithholdingRangeName = "源泉処理";
		private const string InstallmentFirstDueDateRangeName = "第１回期限";
		private const string InstallmentAmountRangeName = "分割金";
		private const string InstallmentChangeRoundRangeName = "変更回";
		private const string InstallmentChangedAmountRangeName = "変更後分割金";
		private const string InstallmentDepositAmountRangeName = "お預かり金額";

		private const string InstallmentPaymentTotalCellAddress = "I9";
		private const string InstallmentMarkerCellAddress = "B12";
		private const string InstallmentIssueDateSourceCellAddress = "J1";
		private const string InstallmentResetMarkerText = "※";
		private const string InstallmentIssueDatePlaceholderText = "発行日";
		private const string DepositAppliedText = "(充当済み)";
		private const string DateDisplayFormat = "yyyy/MM/dd";

		private const string InvoiceBillingSubtotalCellAddress = "F23";
		private const string InvoiceExpenseCellAddress = "F25";
		private const string InvoiceWithholdingFlagCellAddress = "Y24";
		private const string InvoiceFirstDueDateCellAddress = "G10";
		private const string InvoiceDepositAmountCellAddress = "F29";

		private const double BalanceTolerance = 0.5;

		private static readonly string[] RequiredNamedRanges = {
			InstallmentBillingAmountRangeName,
			"お支払い額計",
			InstallmentAmountRangeName,
			InstallmentFirstDueDateRangeName,
			InstallmentExpenseAmountRangeName,
			InstallmentWithholdingRangeName,
			InstallmentDepositAmountRangeName,
			InstallmentChangedAmountRangeName,
			InstallmentChangeRoundRangeName
		};

		private readonly AccountingWorkbookService _accountingWorkbookService;
		private readonly UserErrorService _userErrorService;
		private readonly Logger _logger;

		internal AccountingInstallmentScheduleCommandService (
			AccountingWorkbookService accountingWorkbookService,
			UserErrorService userErrorService,
			Logger logger)
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

			const string procedureName = "AccountingInstallmentSchedule.LoadFormState";
			List<string> warnings = new List<string> ();
			double billingSubtotal = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, InvoiceSheetName, InvoiceBillingSubtotalCellAddress, "請求額小計", procedureName, warnings);
			double expenseAmount = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, InvoiceSheetName, InvoiceExpenseCellAddress, "実費等総額", procedureName, warnings);
			double depositAmount = ReadLoadFormStateDoubleAllowBlankAsZero (workbook, InvoiceSheetName, InvoiceDepositAmountCellAddress, "お預かり金額", procedureName, warnings);
			bool withholding = ReadBooleanSafe (workbook, InvoiceSheetName, InvoiceWithholdingFlagCellAddress);

			AccountingInstallmentScheduleFormState state = new AccountingInstallmentScheduleFormState {
				BillingAmountText = FormatAmount (billingSubtotal + expenseAmount),
				ExpenseAmountText = FormatAmount (expenseAmount),
				WithholdingText = withholding ? "する" : "しない",
				FirstDueDateText = ReadFormattedDateSafe (workbook, InvoiceSheetName, InvoiceFirstDueDateCellAddress),
				DepositAmountText = FormatAmount (depositAmount),
				InstallmentAmountText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, InstallmentAmountRangeName),
				ChangeRoundText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, InstallmentChangeRoundRangeName),
				ChangedInstallmentAmountText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, InstallmentChangedAmountRangeName)
			};

			if (warnings.Count > 0) {
				ShowLoadFormStateNumericReadWarning (string.Join (Environment.NewLine, warnings));
			}

			return state;
		}

		internal AccountingInstallmentScheduleFormState CreateSchedule (Workbook workbook, AccountingInstallmentScheduleCreateRequest request)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}

			if (!TryValidateCreateRequest (request, out DateTime firstDueDate, out double billingAmount, out double expenseAmount, out double depositAmount, out double installmentAmount)) {
				return LoadFormState (workbook);
			}

			ScheduleSnapshot snapshot = null;
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				using (_accountingWorkbookService.BeginInitializationScope ()) {
					SyncInstallmentHeaderFromInvoice (workbook);
					EnsureWorkbookStructure (workbook);
					WriteCreateBaseValues (workbook, request, firstDueDate, billingAmount, expenseAmount, depositAmount, installmentAmount);

					InstallmentScheduleInputs inputs = LoadInputsFromNamedRanges (workbook);
					snapshot = TakeScheduleSnapshot (workbook, AccountingInstallmentSchedulePlanPolicy.StartValueRow);
					ClearScheduleRows (workbook, AccountingInstallmentSchedulePlanPolicy.FirstScheduleRow);
					WriteStartValues (workbook, inputs);

					int lastRow = GenerateSchedule (workbook, inputs, AccountingInstallmentSchedulePlanPolicy.FirstScheduleRow, isCreateFromFirstRow: true);
					FinishSchedule (workbook, lastRow);
					_logger.Info ("Installment schedule created. lastRow=" + lastRow.ToString (CultureInfo.InvariantCulture));
				}
			} catch (Exception exception) {
				RestoreSnapshotQuietly (workbook, snapshot);
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.CreateSchedule", exception);
			} finally {
				ProtectSheetQuietly (workbook, "Installment schedule reprotect failed after create.");
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

			if (!TryValidateChangeRequest (request, out int changeRound, out double changedInstallmentAmount)) {
				return LoadFormState (workbook);
			}

			ScheduleSnapshot snapshot = null;
			try {
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				using (_accountingWorkbookService.BeginInitializationScope ()) {
					EnsureWorkbookStructure (workbook);
					WriteChangeBaseValues (workbook, changeRound, changedInstallmentAmount);

					InstallmentScheduleInputs inputs = LoadInputsFromNamedRanges (workbook);
					List<AccountingInstallmentScheduleExistingRow> existingRows = ReadExistingScheduleRows (workbook);
					AccountingInstallmentScheduleChangeStart changeStart = AccountingInstallmentSchedulePlanPolicy.ResolveChangeStart (existingRows, changeRound);

					snapshot = TakeScheduleSnapshot (workbook, changeStart.StartRow);
					ClearScheduleRows (workbook, changeStart.StartRow);

					int lastRow = GenerateSchedule (workbook, inputs, changeStart.StartRow, isCreateFromFirstRow: false);
					FinishSchedule (workbook, lastRow);
					_logger.Info ("Installment schedule change applied. startRow=" + changeStart.StartRow.ToString (CultureInfo.InvariantCulture) + ", lastRow=" + lastRow.ToString (CultureInfo.InvariantCulture));
				}
			} catch (Exception exception) {
				RestoreSnapshotQuietly (workbook, snapshot);
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.ApplyChange", exception);
			} finally {
				ProtectSheetQuietly (workbook, "Installment schedule reprotect failed after change.");
			}

			return LoadFormState (workbook);
		}

		internal AccountingInstallmentScheduleFormState ApplyIssueDate (Workbook workbook)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (workbook, SheetName, InstallmentIssueDateSourceCellAddress);
				_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentIssueDateRangeName, value);
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
				_accountingWorkbookService.UnprotectSheet (workbook, SheetName);
				using (_accountingWorkbookService.BeginInitializationScope ()) {
					SyncInstallmentHeaderFromInvoice (workbook);
					_accountingWorkbookService.WriteCell (workbook, SheetName, InstallmentMarkerCellAddress, InstallmentResetMarkerText);
					_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentIssueDateRangeName, InstallmentIssueDatePlaceholderText);
					_accountingWorkbookService.ClearNamedRangeMergeAreaContents (workbook, SheetName, InstallmentBillingAmountRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentAmountRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentFirstDueDateRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentExpenseAmountRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentWithholdingRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentChangeRoundRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentChangedAmountRangeName);
					_accountingWorkbookService.ClearNamedRangeContents (workbook, SheetName, InstallmentDepositAmountRangeName);
					ClearScheduleRows (workbook, AccountingInstallmentSchedulePlanPolicy.StartValueRow);
					_accountingWorkbookService.ClearRangeContents (workbook, SheetName, InstallmentPaymentTotalCellAddress);
					_accountingWorkbookService.SetPrintAreaByBounds (workbook, SheetName, AccountingInstallmentSchedulePlanPolicy.StartValueRow, AccountingInstallmentSchedulePlanPolicy.PrintLastColumn);
					EnsureHiddenLayout (workbook);
				}
			} catch (Exception exception) {
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.Reset", exception);
			} finally {
				ProtectSheetQuietly (workbook, "Installment schedule reprotect failed.");
			}

			return LoadFormState (workbook);
		}

		private bool TryValidateCreateRequest (
			AccountingInstallmentScheduleCreateRequest request,
			out DateTime firstDueDate,
			out double billingAmount,
			out double expenseAmount,
			out double depositAmount,
			out double installmentAmount)
		{
			firstDueDate = DateTime.MinValue;
			billingAmount = 0;
			expenseAmount = 0;
			depositAmount = 0;
			installmentAmount = 0;

			if (!TryParseAmountInput (request.BillingAmountText, "請求額", allowBlankAsZero: false, out billingAmount) ||
				!TryParseAmountInput (request.ExpenseAmountText, "実費等総額", allowBlankAsZero: true, out expenseAmount) ||
				!TryParseAmountInput (request.DepositAmountText, "お預かり金額", allowBlankAsZero: true, out depositAmount) ||
				!TryParseAmountInput (request.InstallmentAmountText, "分割金", allowBlankAsZero: false, out installmentAmount)) {
				return false;
			}

			if (billingAmount <= 0) {
				UserErrorService.ShowOkNotification ("請求額が0円です。請求書を作成してください。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (expenseAmount < 0 || depositAmount < 0) {
				UserErrorService.ShowOkNotification ("実費等総額・お預かり金額に負数は指定できません。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (depositAmount > billingAmount) {
				UserErrorService.ShowOkNotification ("お預かり金額が請求額を超えています。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (installmentAmount <= 0) {
				UserErrorService.ShowOkNotification ("分割払い額が未入力です。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (!TryParseDateInput (request.FirstDueDateText, out firstDueDate)) {
				UserErrorService.ShowOkNotification ("期限（1回目）欄に日付が正しく入力されていません\r\n請求書で期限を入力してください", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}

			return true;
		}

		private bool TryValidateChangeRequest (
			AccountingInstallmentScheduleChangeRequest request,
			out int changeRound,
			out double changedInstallmentAmount)
		{
			changeRound = 0;
			changedInstallmentAmount = 0;

			if (!TryParseAmountInput (request.ChangedInstallmentAmountText, "変更後分割金", allowBlankAsZero: false, out changedInstallmentAmount)) {
				return false;
			}

			if (changedInstallmentAmount <= 0) {
				UserErrorService.ShowOkNotification ("変更後の分割払い額が未入力です。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (!TryParseIntegerInput (request.ChangeRoundText, "変更回", out changeRound) || changeRound < 1) {
				UserErrorService.ShowOkNotification ("変更回は 1 以上の整数で入力してください。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}

			return true;
		}

		private void WriteCreateBaseValues (
			Workbook workbook,
			AccountingInstallmentScheduleCreateRequest request,
			DateTime firstDueDate,
			double billingAmount,
			double expenseAmount,
			double depositAmount,
			double installmentAmount)
		{
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentBillingAmountRangeName, billingAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentExpenseAmountRangeName, expenseAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentWithholdingRangeName, request.WithholdingText ?? string.Empty);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentFirstDueDateRangeName, firstDueDate);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentAmountRangeName, installmentAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentDepositAmountRangeName, depositAmount);
		}

		private void WriteChangeBaseValues (Workbook workbook, int changeRound, double changedInstallmentAmount)
		{
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentChangeRoundRangeName, changeRound);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentChangedAmountRangeName, changedInstallmentAmount);
			_accountingWorkbookService.WriteNamedRangeValue (workbook, SheetName, InstallmentAmountRangeName, changedInstallmentAmount);
		}

		private InstallmentScheduleInputs LoadInputsFromNamedRanges (Workbook workbook)
		{
			EnsureRequiredNamedRanges (workbook);

			double billingAmount = ReadRequiredNamedDouble (workbook, InstallmentBillingAmountRangeName, "請求額");
			double expenseAmount = ReadRequiredNamedDouble (workbook, InstallmentExpenseAmountRangeName, "実費等総額");
			double depositAmount = ReadRequiredNamedDouble (workbook, InstallmentDepositAmountRangeName, "お預かり金額");
			double installmentAmount = ReadRequiredNamedDouble (workbook, InstallmentAmountRangeName, "分割金");
			DateTime firstDueDate = ReadRequiredNamedDate (workbook, InstallmentFirstDueDateRangeName, "第１回期限");
			string withholdingText = ReadRequiredNamedText (workbook, InstallmentWithholdingRangeName, "源泉処理");

			if (billingAmount <= 0) {
				throw new InvalidOperationException ("請求額は 1 円以上で入力してください。");
			}
			if (expenseAmount < 0 || depositAmount < 0) {
				throw new InvalidOperationException ("実費等総額・お預かり金額に負数は指定できません。");
			}
			if (depositAmount > billingAmount) {
				throw new InvalidOperationException ("お預かり金額が請求額を超えています。");
			}
			if (installmentAmount <= 0) {
				throw new InvalidOperationException ("分割金は 1 円以上で入力してください。");
			}
			if (!IsSupportedWithholdingText (withholdingText)) {
				throw new InvalidOperationException ("源泉処理は「する」または「しない」を指定してください。");
			}

			return new InstallmentScheduleInputs (
				billingAmount,
				expenseAmount,
				depositAmount,
				installmentAmount,
				firstDueDate);
		}

		private int GenerateSchedule (Workbook workbook, InstallmentScheduleInputs inputs, int startRow, bool isCreateFromFirstRow)
		{
			int row = startRow;
			while (true) {
				AccountingInstallmentSchedulePlanPolicy.EnsureWritableRow (row);
				ScheduleRowGoal rowGoal = WriteScheduleRow (workbook, inputs, row, isCreateFromFirstRow && row == AccountingInstallmentSchedulePlanPolicy.FirstScheduleRow);
				GoalSeekAndVerify (workbook, row, rowGoal.FormulaColumn, rowGoal.TargetValue, rowGoal.Kind);

				double billingBalance = ReadRequiredDouble (workbook, "I" + row.ToString (CultureInfo.InvariantCulture), "請求残高", "AccountingInstallmentSchedule.GenerateSchedule");
				if (billingBalance < -BalanceTolerance) {
					GoalSeekAndVerify (workbook, row, "I", 0, "最終回請求残高");
					return row;
				}
				if (Math.Abs (billingBalance) <= BalanceTolerance) {
					if (Math.Abs (billingBalance) > 0.0001) {
						GoalSeekAndVerify (workbook, row, "I", 0, "最終回請求残高");
					}
					return row;
				}

				row++;
				if (row > AccountingInstallmentSchedulePlanPolicy.LastScheduleRow) {
					throw new InvalidOperationException ("分割払い予定表が A12:J73 のテーブル範囲を超えます。分割金を増額してください。");
				}
			}
		}

		private ScheduleRowGoal WriteScheduleRow (Workbook workbook, InstallmentScheduleInputs inputs, int row, bool isCreateFirstRow)
		{
			int previousRow = AccountingInstallmentSchedulePlanPolicy.GetPreviousBalanceRowForDetailRow (row);

			if (isCreateFirstRow) {
				bool hasDeposit = inputs.DepositAmount != 0;
				_accountingWorkbookService.WriteCellValue (workbook, SheetName, "A" + row.ToString (CultureInfo.InvariantCulture), hasDeposit ? 0 : 1);
				if (hasDeposit) {
					_accountingWorkbookService.WriteCell (workbook, SheetName, "B" + row.ToString (CultureInfo.InvariantCulture), DepositAppliedText);
					_accountingWorkbookService.SetHorizontalAlignmentCenter (workbook, SheetName, "B" + row.ToString (CultureInfo.InvariantCulture));
				} else {
					_accountingWorkbookService.WriteCellValue (workbook, SheetName, "B" + row.ToString (CultureInfo.InvariantCulture), inputs.FirstDueDate);
				}

				double firstExpenseCharge = AccountingInstallmentSchedulePlanPolicy.ResolveFirstRowExpenseCharge (inputs.DepositAmount, inputs.InstallmentAmount, inputs.ExpenseAmount);
				WriteScheduleRowFormulas (workbook, row, previousRow, firstExpenseCharge);
				double target = hasDeposit ? inputs.DepositAmount : inputs.InstallmentAmount;
				return new ScheduleRowGoal ("C", target, hasDeposit ? "14行目お預かり金額" : "14行目分割金");
			}

			double previousExpenseBalance = ReadRequiredDouble (workbook, "J" + previousRow.ToString (CultureInfo.InvariantCulture), "実費残高", "AccountingInstallmentSchedule.WriteScheduleRow");
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "A" + row.ToString (CultureInfo.InvariantCulture), "=A" + previousRow.ToString (CultureInfo.InvariantCulture) + "+1");
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, "B" + row.ToString (CultureInfo.InvariantCulture), ResolveDueDate (workbook, inputs, previousRow));
			double expenseCharge = AccountingInstallmentSchedulePlanPolicy.ResolveExpenseCharge (inputs.InstallmentAmount, previousExpenseBalance);
			WriteScheduleRowFormulas (workbook, row, previousRow, expenseCharge);
			return new ScheduleRowGoal ("C", inputs.InstallmentAmount, "通常分割金");
		}

		private void WriteScheduleRowFormulas (Workbook workbook, int row, int previousRow, double expenseCharge)
		{
			string rowText = row.ToString (CultureInfo.InvariantCulture);
			string previousRowText = previousRow.ToString (CultureInfo.InvariantCulture);

			_accountingWorkbookService.WriteCellValue (workbook, SheetName, "D" + rowText, 0);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, "H" + rowText, expenseCharge);
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "C" + rowText, "=F" + rowText + "-G" + rowText + "+H" + rowText);
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "E" + rowText, "=ROUNDDOWN(D" + rowText + "*引数!$C$4,0)");
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "F" + rowText, "=D" + rowText + "+E" + rowText);
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "G" + rowText, "=IF($J$9,ROUNDDOWN(IF(D" + rowText + "<=1000000,D" + rowText + "*10.21%,(D" + rowText + "-1000000)*20.42%+102100),0),0)");
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "I" + rowText, "=I" + previousRowText + "-(C" + rowText + "+G" + rowText + ")");
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "J" + rowText, "=J" + previousRowText + "-H" + rowText);
		}

		private DateTime ResolveDueDate (Workbook workbook, InstallmentScheduleInputs inputs, int previousRow)
		{
			string previousDueDateAddress = "B" + previousRow.ToString (CultureInfo.InvariantCulture);
			if (TryReadCellDate (workbook, previousDueDateAddress, out DateTime previousDueDate)) {
				return AccountingInstallmentSchedulePlanPolicy.ResolveNextMonthEndDueDate (previousDueDate);
			}

			return inputs.FirstDueDate;
		}

		private void GoalSeekAndVerify (Workbook workbook, int row, string formulaColumn, double targetValue, string kind)
		{
			string rowText = row.ToString (CultureInfo.InvariantCulture);
			string formulaCellAddress = formulaColumn + rowText;
			string changingCellAddress = "D" + rowText;

			_accountingWorkbookService.ExecuteGoalSeekOrThrow (workbook, SheetName, formulaCellAddress, changingCellAddress, targetValue);

			double current = ReadRequiredDouble (workbook, formulaCellAddress, kind, "AccountingInstallmentSchedule.GoalSeekAndVerify");
			double changingValue = ReadRequiredDouble (workbook, changingCellAddress, "10％対象計", "AccountingInstallmentSchedule.GoalSeekAndVerify");
			if (Math.Abs (current - targetValue) > BalanceTolerance) {
				throw new InvalidOperationException (
					"GoalSeek 後の値が目標値と一致しません。row=" + rowText +
					", formulaCell=" + formulaCellAddress +
					", changingCell=" + changingCellAddress +
					", target=" + targetValue.ToString (CultureInfo.InvariantCulture) +
					", current=" + current.ToString (CultureInfo.InvariantCulture));
			}
			if (changingValue < -BalanceTolerance) {
				throw new InvalidOperationException (
					"GoalSeek 後の10％対象計が負数です。row=" + rowText +
					", formulaCell=" + formulaCellAddress +
					", changingCell=" + changingCellAddress +
					", target=" + targetValue.ToString (CultureInfo.InvariantCulture) +
					", changingValue=" + changingValue.ToString (CultureInfo.InvariantCulture));
			}
		}

		private void WriteStartValues (Workbook workbook, InstallmentScheduleInputs inputs)
		{
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, "I13", inputs.BillingAmount);
			_accountingWorkbookService.WriteCellValue (workbook, SheetName, "J13", inputs.ExpenseAmount);
			RestoreWithholdingHelperFormula (workbook);
			EnsureHiddenLayout (workbook);
		}

		private void FinishSchedule (Workbook workbook, int lastRow)
		{
			ClearScheduleRows (workbook, lastRow + 1);
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, InstallmentPaymentTotalCellAddress, "=SUM(C14:C" + lastRow.ToString (CultureInfo.InvariantCulture) + ")");
			_accountingWorkbookService.SetPrintAreaByBounds (workbook, SheetName, lastRow, AccountingInstallmentSchedulePlanPolicy.PrintLastColumn);
			EnsureHiddenLayout (workbook);
		}

		private void EnsureWorkbookStructure (Workbook workbook)
		{
			EnsureRequiredNamedRanges (workbook);
			if (!_accountingWorkbookService.HasListObjectRange (workbook, SheetName, TableRangeAddress)) {
				throw new InvalidOperationException ("分割払い予定表シートのテーブル範囲が A12:J73 ではありません。実ブック構造を確認してください。");
			}
			RestoreWithholdingHelperFormula (workbook);
			EnsureHiddenLayout (workbook);
		}

		private void EnsureRequiredNamedRanges (Workbook workbook)
		{
			for (int index = 0; index < RequiredNamedRanges.Length; index++) {
				_accountingWorkbookService.GetNamedRangeAddress (workbook, SheetName, RequiredNamedRanges [index]);
			}
		}

		private void RestoreWithholdingHelperFormula (Workbook workbook)
		{
			_accountingWorkbookService.WriteCellFormula (workbook, SheetName, "J9", "=IF(K9=\"する\",TRUE,FALSE)");
		}

		private void EnsureHiddenLayout (Workbook workbook)
		{
			_accountingWorkbookService.SetRowHidden (workbook, SheetName, AccountingInstallmentSchedulePlanPolicy.StartValueRow, true);
			_accountingWorkbookService.SetColumnHidden (workbook, SheetName, "J", true);
		}

		private List<AccountingInstallmentScheduleExistingRow> ReadExistingScheduleRows (Workbook workbook)
		{
			List<AccountingInstallmentScheduleExistingRow> rows = new List<AccountingInstallmentScheduleExistingRow> ();
			for (int row = AccountingInstallmentSchedulePlanPolicy.FirstScheduleRow; row <= AccountingInstallmentSchedulePlanPolicy.LastScheduleRow; row++) {
				if (!TryReadScheduleRound (workbook, row, out int round, out bool isBlank)) {
					if (isBlank) {
						string residualCellAddress = FindFirstResidualScheduleCellAfterTerminator (workbook, row);
						AccountingInstallmentSchedulePlanPolicy.EnsureNoExistingScheduleContentAfterTerminator (row, residualCellAddress);
						break;
					}

					continue;
				}
				double billingBalance = TryReadDouble (workbook, "I" + row.ToString (CultureInfo.InvariantCulture), out double billing) ? billing : double.NaN;
				double expenseBalance = TryReadDouble (workbook, "J" + row.ToString (CultureInfo.InvariantCulture), out double expense) ? expense : double.NaN;
				rows.Add (new AccountingInstallmentScheduleExistingRow (row, round, billingBalance, expenseBalance));
			}

			return rows;
		}

		private bool TryReadScheduleRound (Workbook workbook, int row, out int round, out bool isBlank)
		{
			round = 0;
			isBlank = false;
			string address = "A" + row.ToString (CultureInfo.InvariantCulture);
			object value = _accountingWorkbookService.ReadCellValue (workbook, SheetName, address);
			string displayText = _accountingWorkbookService.ReadDisplayText (workbook, SheetName, address);
			if (!AccountingNumericCellReader.TryParseNumericCell (value, displayText, out double parsed, out isBlank)) {
				if (isBlank) {
					return false;
				}

				throw AccountingNumericCellReader.CreateReadFailureException (SheetName, address, "回数", "AccountingInstallmentSchedule.ReadExistingScheduleRows", displayText, allowBlankAsZero: false);
			}
			if (Math.Abs (parsed - Math.Round (parsed)) > 0.0001) {
				throw new InvalidOperationException (address + " の回数が整数ではありません。");
			}
			round = Convert.ToInt32 (Math.Round (parsed), CultureInfo.InvariantCulture);
			return true;
		}

		private string FindFirstResidualScheduleCellAfterTerminator (Workbook workbook, int terminatorRow)
		{
			string sameRowResidual = FindFirstNonBlankScheduleCell (workbook, terminatorRow, 2);
			if (!string.IsNullOrWhiteSpace (sameRowResidual)) {
				return sameRowResidual;
			}

			return FindFirstNonBlankScheduleCell (workbook, terminatorRow + 1, 1);
		}

		private string FindFirstNonBlankScheduleCell (Workbook workbook, int startRow, int startColumn)
		{
			if (startRow > AccountingInstallmentSchedulePlanPolicy.LastScheduleRow) {
				return null;
			}

			int firstColumn = Math.Max (1, startColumn);
			for (int row = startRow; row <= AccountingInstallmentSchedulePlanPolicy.LastScheduleRow; row++) {
				for (int column = firstColumn; column <= AccountingInstallmentSchedulePlanPolicy.LastScheduleColumn; column++) {
					string address = GetScheduleCellAddress (row, column);
					object value = _accountingWorkbookService.ReadCellValue (workbook, SheetName, address);
					if (!IsBlankCellValue (value)) {
						return address;
					}
				}
				firstColumn = 1;
			}

			return null;
		}

		private static string GetScheduleCellAddress (int row, int column)
		{
			char columnName = (char)('A' + column - 1);
			return columnName.ToString () + row.ToString (CultureInfo.InvariantCulture);
		}

		private static bool IsBlankCellValue (object value)
		{
			if (value == null) {
				return true;
			}

			string text = value as string;
			return text != null && string.IsNullOrWhiteSpace (text);
		}

		private ScheduleSnapshot TakeScheduleSnapshot (Workbook workbook, int startRow)
		{
			if (startRow > AccountingInstallmentSchedulePlanPolicy.LastScheduleRow) {
				return null;
			}

			string address = "A" + startRow.ToString (CultureInfo.InvariantCulture) + ":J" + AccountingInstallmentSchedulePlanPolicy.LastScheduleRow.ToString (CultureInfo.InvariantCulture);
			object rangeSnapshot = _accountingWorkbookService.ReadRangeFormulaSnapshot (workbook, SheetName, address);
			object totalSnapshot = _accountingWorkbookService.ReadRangeFormulaSnapshot (workbook, SheetName, InstallmentPaymentTotalCellAddress);
			return new ScheduleSnapshot (address, rangeSnapshot, totalSnapshot);
		}

		private void RestoreSnapshotQuietly (Workbook workbook, ScheduleSnapshot snapshot)
		{
			if (workbook == null || snapshot == null) {
				return;
			}

			try {
				_accountingWorkbookService.WriteRangeFormulaSnapshot (workbook, SheetName, snapshot.Address, snapshot.RangeFormulaSnapshot);
				_accountingWorkbookService.WriteRangeFormulaSnapshot (workbook, SheetName, InstallmentPaymentTotalCellAddress, snapshot.PaymentTotalFormulaSnapshot);
			} catch (Exception exception) {
				_logger.Error ("Installment schedule snapshot restore failed.", exception);
			}
		}

		private void ClearScheduleRows (Workbook workbook, int startRow)
		{
			if (startRow > AccountingInstallmentSchedulePlanPolicy.LastScheduleRow) {
				return;
			}

			if (startRow < AccountingInstallmentSchedulePlanPolicy.StartValueRow) {
				startRow = AccountingInstallmentSchedulePlanPolicy.StartValueRow;
			}

			string range = "A" + startRow.ToString (CultureInfo.InvariantCulture) + ":J" + AccountingInstallmentSchedulePlanPolicy.LastScheduleRow.ToString (CultureInfo.InvariantCulture);
			_accountingWorkbookService.ClearRangeContents (workbook, SheetName, range);
		}

		private double ReadRequiredNamedDouble (Workbook workbook, string rangeName, string itemName)
		{
			object cellValue = _accountingWorkbookService.ReadNamedRangeValue (workbook, SheetName, rangeName);
			string displayText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, rangeName);
			if (AccountingNumericCellReader.TryParseNumericCell (cellValue, displayText, out double value, out bool isBlank) && !isBlank) {
				return value;
			}

			string address = _accountingWorkbookService.GetNamedRangeAddress (workbook, SheetName, rangeName);
			throw AccountingNumericCellReader.CreateReadFailureException (SheetName, address, itemName, "AccountingInstallmentSchedule.LoadInputs", displayText, allowBlankAsZero: false);
		}

		private DateTime ReadRequiredNamedDate (Workbook workbook, string rangeName, string itemName)
		{
			object value = _accountingWorkbookService.ReadNamedRangeValue (workbook, SheetName, rangeName);
			if (value is double) {
				return DateTime.FromOADate ((double)value);
			}
			if (value is DateTime) {
				return (DateTime)value;
			}

			string displayText = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, rangeName);
			if (TryParseDateInput (displayText, out DateTime parsed)) {
				return parsed;
			}

			throw new InvalidOperationException (itemName + " に日付が入力されていません。");
		}

		private string ReadRequiredNamedText (Workbook workbook, string rangeName, string itemName)
		{
			string text = _accountingWorkbookService.ReadDisplayTextByNamedRange (workbook, SheetName, rangeName);
			if (string.IsNullOrWhiteSpace (text)) {
				text = Convert.ToString (_accountingWorkbookService.ReadNamedRangeValue (workbook, SheetName, rangeName), CultureInfo.InvariantCulture) ?? string.Empty;
			}
			if (string.IsNullOrWhiteSpace (text)) {
				throw new InvalidOperationException (itemName + " が未入力です。");
			}
			return text.Trim ();
		}

		private bool TryReadDouble (Workbook workbook, string address, out double value)
		{
			value = 0;
			object cellValue = _accountingWorkbookService.ReadCellValue (workbook, SheetName, address);
			string displayText = _accountingWorkbookService.ReadDisplayText (workbook, SheetName, address);
			return AccountingNumericCellReader.TryParseNumericCell (cellValue, displayText, out value, out bool isBlank) && !isBlank;
		}

		private double ReadRequiredDouble (Workbook workbook, string address, string itemName, string procedureName)
		{
			return ReadRequiredDouble (workbook, SheetName, address, itemName, procedureName);
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
				return 0;
			}
		}

		private double ReadNumericCellCore (Workbook workbook, string sheetName, string address, string itemName, string procedureName, bool allowBlankAsZero)
		{
			object cellValue = _accountingWorkbookService.ReadCellValue (workbook, sheetName, address);
			string displayText = _accountingWorkbookService.ReadDisplayText (workbook, sheetName, address);
			if (AccountingNumericCellReader.TryParseNumericCell (cellValue, displayText, out double value, out bool isBlank)) {
				return value;
			}
			if (allowBlankAsZero && isBlank) {
				return 0;
			}

			InvalidOperationException ex = AccountingNumericCellReader.CreateReadFailureException (sheetName, address, itemName, procedureName, displayText, allowBlankAsZero);
			string valueText = Convert.ToString (cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
			_logger.Error ("Accounting numeric cell read failed. sheet=" + sheetName + ", address=" + address + ", item=" + itemName + ", procedure=" + procedureName + ", displayText=" + (string.IsNullOrWhiteSpace (displayText) ? "（空欄）" : displayText.Trim ()) + ", cellValue=" + valueText, ex);
			throw ex;
		}

		private bool TryReadCellDate (Workbook workbook, string address, out DateTime value)
		{
			try {
				value = _accountingWorkbookService.ReadDateCell (workbook, SheetName, address);
				return true;
			} catch {
				value = DateTime.MinValue;
				return false;
			}
		}

		private void SyncInstallmentHeaderFromInvoice (Workbook workbook)
		{
			_accountingWorkbookService.CopyValueRange (workbook, InvoiceSheetName, "A3:A4", SheetName, "A3:A4");
		}

		private void ProtectSheetQuietly (Workbook workbook, string logMessage)
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

			_logger.Warn ("AccountingInstallmentSchedule.LoadFormState numeric read warning. " + warningMessage.Replace (Environment.NewLine, " | "));
			UserErrorService.ShowOkNotification ("数値読取に失敗した項目があります。該当項目は 0 として表示しています。" + Environment.NewLine + Environment.NewLine + warningMessage, "案件情報System", MessageBoxIcon.Warning);
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

		private static bool TryParseAmountInput (string text, string itemName, bool allowBlankAsZero, out double value)
		{
			value = 0;
			string normalized = (text ?? string.Empty).Replace (",", string.Empty).Trim ();
			if (normalized.Length == 0) {
				if (allowBlankAsZero) {
					return true;
				}
				UserErrorService.ShowOkNotification (itemName + " が未入力です。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (!double.TryParse (normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out value) &&
				!double.TryParse (normalized, NumberStyles.Number, CultureInfo.CurrentCulture, out value)) {
				UserErrorService.ShowOkNotification (itemName + " は数値で入力してください。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			return true;
		}

		private static bool TryParseIntegerInput (string text, string itemName, out int value)
		{
			value = 0;
			string normalized = (text ?? string.Empty).Replace (",", string.Empty).Trim ();
			if (normalized.Length == 0) {
				UserErrorService.ShowOkNotification (itemName + " が未入力です。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			if (!int.TryParse (normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) &&
				!int.TryParse (normalized, NumberStyles.Integer, CultureInfo.CurrentCulture, out value)) {
				UserErrorService.ShowOkNotification (itemName + " は整数で入力してください。", "案件情報System", MessageBoxIcon.Asterisk);
				return false;
			}
			return true;
		}

		private static bool TryParseDateInput (string text, out DateTime value)
		{
			return DateTime.TryParse (text, CultureInfo.InvariantCulture, DateTimeStyles.None, out value) ||
				DateTime.TryParse (text, CultureInfo.CurrentCulture, DateTimeStyles.None, out value);
		}

		private static bool IsSupportedWithholdingText (string text)
		{
			return string.Equals ((text ?? string.Empty).Trim (), "する", StringComparison.Ordinal) ||
				string.Equals ((text ?? string.Empty).Trim (), "しない", StringComparison.Ordinal);
		}

		private static string FormatAmount (double value)
		{
			return value.ToString ("#,##0", CultureInfo.InvariantCulture);
		}

		private string ReadFormattedDateSafe (Workbook workbook, string sheetName, string address)
		{
			try {
				return _accountingWorkbookService.ReadDateCell (workbook, sheetName, address).ToString (DateDisplayFormat, CultureInfo.InvariantCulture);
			} catch {
				return _accountingWorkbookService.ReadDisplayText (workbook, sheetName, address);
			}
		}

		private sealed class InstallmentScheduleInputs
		{
			internal InstallmentScheduleInputs (
				double billingAmount,
				double expenseAmount,
				double depositAmount,
				double installmentAmount,
				DateTime firstDueDate)
			{
				BillingAmount = billingAmount;
				ExpenseAmount = expenseAmount;
				DepositAmount = depositAmount;
				InstallmentAmount = installmentAmount;
				FirstDueDate = firstDueDate;
			}

			internal double BillingAmount { get; private set; }
			internal double ExpenseAmount { get; private set; }
			internal double DepositAmount { get; private set; }
			internal double InstallmentAmount { get; private set; }
			internal DateTime FirstDueDate { get; private set; }
		}

		private sealed class ScheduleRowGoal
		{
			internal ScheduleRowGoal (string formulaColumn, double targetValue, string kind)
			{
				FormulaColumn = formulaColumn;
				TargetValue = targetValue;
				Kind = kind;
			}

			internal string FormulaColumn { get; private set; }
			internal double TargetValue { get; private set; }
			internal string Kind { get; private set; }
		}

		private sealed class ScheduleSnapshot
		{
			internal ScheduleSnapshot (string address, object rangeFormulaSnapshot, object paymentTotalFormulaSnapshot)
			{
				Address = address;
				RangeFormulaSnapshot = rangeFormulaSnapshot;
				PaymentTotalFormulaSnapshot = paymentTotalFormulaSnapshot;
			}

			internal string Address { get; private set; }
			internal object RangeFormulaSnapshot { get; private set; }
			internal object PaymentTotalFormulaSnapshot { get; private set; }
		}
	}
}
