using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class AccountingInstallmentSchedulePlanPolicyTests
	{
		[Fact]
		public void ResolveNextMonthEndDueDate_UsesFollowingMonthEnd ()
		{
			DateTime result = AccountingInstallmentSchedulePlanPolicy.ResolveNextMonthEndDueDate (new DateTime (2026, 1, 15));

			Assert.Equal (new DateTime (2026, 2, 28), result);
		}

		[Fact]
		public void ResolveExpenseCharge_UsesActivePaymentBasisAndRemainingExpense ()
		{
			Assert.Equal (50000, AccountingInstallmentSchedulePlanPolicy.ResolveExpenseCharge (50000, 120000));
			Assert.Equal (12000, AccountingInstallmentSchedulePlanPolicy.ResolveExpenseCharge (50000, 12000));
			Assert.Equal (0, AccountingInstallmentSchedulePlanPolicy.ResolveExpenseCharge (0, 12000));
		}

		[Fact]
		public void ResolveFirstRowExpenseCharge_WithoutDeposit_UsesInstallmentAmount ()
		{
			Assert.Equal (50000, AccountingInstallmentSchedulePlanPolicy.ResolveFirstRowExpenseCharge (0, 50000, 120000));
			Assert.Equal (12000, AccountingInstallmentSchedulePlanPolicy.ResolveFirstRowExpenseCharge (0, 50000, 12000));
		}

		[Fact]
		public void ResolveFirstRowExpenseCharge_WithDeposit_UsesDepositAmount ()
		{
			Assert.Equal (30000, AccountingInstallmentSchedulePlanPolicy.ResolveFirstRowExpenseCharge (30000, 50000, 120000));
			Assert.Equal (12000, AccountingInstallmentSchedulePlanPolicy.ResolveFirstRowExpenseCharge (30000, 50000, 12000));
		}

		[Fact]
		public void EnsureWritableRow_AllowsRowsInsideA12J73TableDataArea ()
		{
			AccountingInstallmentSchedulePlanPolicy.EnsureWritableRow (14);
			AccountingInstallmentSchedulePlanPolicy.EnsureWritableRow (73);

			Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.EnsureWritableRow (12));
			Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.EnsureWritableRow (13));
			Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.EnsureWritableRow (74));
		}

		[Fact]
		public void ScheduleRows_TreatRow12AsHeaderRow13AsStartValueAndRow14AsFirstDetail ()
		{
			Assert.Equal (12, AccountingInstallmentSchedulePlanPolicy.HeaderRow);
			Assert.Equal (13, AccountingInstallmentSchedulePlanPolicy.StartValueRow);
			Assert.Equal (14, AccountingInstallmentSchedulePlanPolicy.FirstScheduleRow);
			Assert.False (AccountingInstallmentSchedulePlanPolicy.IsScheduleDetailRow (12));
			Assert.False (AccountingInstallmentSchedulePlanPolicy.IsScheduleDetailRow (13));
			Assert.True (AccountingInstallmentSchedulePlanPolicy.IsScheduleDetailRow (14));
		}

		[Fact]
		public void FirstDetailRow_UsesRow13AsPreviousBalanceRow ()
		{
			int previousRow = AccountingInstallmentSchedulePlanPolicy.GetPreviousBalanceRowForDetailRow (AccountingInstallmentSchedulePlanPolicy.FirstScheduleRow);

			Assert.Equal (AccountingInstallmentSchedulePlanPolicy.StartValueRow, previousRow);
			Assert.NotEqual (AccountingInstallmentSchedulePlanPolicy.HeaderRow, previousRow);
		}

		[Fact]
		public void ResolveChangeStart_SelectsMatchingChangeRound ()
		{
			List<AccountingInstallmentScheduleExistingRow> rows = new List<AccountingInstallmentScheduleExistingRow> {
				new AccountingInstallmentScheduleExistingRow (14, 0, 90000, 30000),
				new AccountingInstallmentScheduleExistingRow (15, 1, 60000, 10000),
				new AccountingInstallmentScheduleExistingRow (16, 2, 30000, 0)
			};

			AccountingInstallmentScheduleChangeStart result = AccountingInstallmentSchedulePlanPolicy.ResolveChangeStart (rows, 2);

			Assert.Equal (16, result.StartRow);
			Assert.Equal (15, result.PreviousRow);
		}

		[Fact]
		public void ResolveChangeStart_RejectsStartRowChange ()
		{
			List<AccountingInstallmentScheduleExistingRow> rows = new List<AccountingInstallmentScheduleExistingRow> {
				new AccountingInstallmentScheduleExistingRow (14, 1, 90000, 30000),
				new AccountingInstallmentScheduleExistingRow (15, 2, 60000, 10000)
			};

			Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.ResolveChangeStart (rows, 1));
		}

		[Fact]
		public void ResolveChangeStart_RejectsMissingChangeRound ()
		{
			List<AccountingInstallmentScheduleExistingRow> rows = new List<AccountingInstallmentScheduleExistingRow> {
				new AccountingInstallmentScheduleExistingRow (14, 0, 90000, 30000),
				new AccountingInstallmentScheduleExistingRow (15, 1, 60000, 10000)
			};

			InvalidOperationException exception = Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.ResolveChangeStart (rows, 5));

			Assert.Contains ("変更回 5", exception.Message);
			Assert.Contains ("既存予定表", exception.Message);
		}

		[Fact]
		public void ResolveChangeStart_RejectsSkippedChangeRound ()
		{
			List<AccountingInstallmentScheduleExistingRow> rows = new List<AccountingInstallmentScheduleExistingRow> {
				new AccountingInstallmentScheduleExistingRow (14, 1, 90000, 30000),
				new AccountingInstallmentScheduleExistingRow (15, 2, 60000, 10000),
				new AccountingInstallmentScheduleExistingRow (16, 4, 30000, 0)
			};

			InvalidOperationException exception = Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.ResolveChangeStart (rows, 3));

			Assert.Contains ("変更回 3", exception.Message);
			Assert.Contains ("見つかりません", exception.Message);
		}

		[Fact]
		public void EnsureNoExistingScheduleContentAfterTerminator_AllowsBlankUnusedRows ()
		{
			AccountingInstallmentSchedulePlanPolicy.EnsureNoExistingScheduleContentAfterTerminator (23, null);
			AccountingInstallmentSchedulePlanPolicy.EnsureNoExistingScheduleContentAfterTerminator (23, string.Empty);
		}

		[Fact]
		public void EnsureNoExistingScheduleContentAfterTerminator_RejectsResidualCellsAfterBlankRound ()
		{
			InvalidOperationException exception = Assert.Throws<InvalidOperationException> (
				() => AccountingInstallmentSchedulePlanPolicy.EnsureNoExistingScheduleContentAfterTerminator (23, "A24"));

			Assert.Contains ("空欄行", exception.Message);
			Assert.Contains ("A23", exception.Message);
			Assert.Contains ("A24", exception.Message);
		}

		[Fact]
		public void ResolveChangeStart_RejectsBrokenPreviousBalance ()
		{
			List<AccountingInstallmentScheduleExistingRow> rows = new List<AccountingInstallmentScheduleExistingRow> {
				new AccountingInstallmentScheduleExistingRow (14, 0, double.NaN, 30000),
				new AccountingInstallmentScheduleExistingRow (15, 1, 60000, 10000)
			};

			Assert.Throws<InvalidOperationException> (() => AccountingInstallmentSchedulePlanPolicy.ResolveChangeStart (rows, 1));
		}
	}
}
