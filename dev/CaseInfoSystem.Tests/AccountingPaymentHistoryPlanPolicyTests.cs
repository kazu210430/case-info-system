using System;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class AccountingPaymentHistoryPlanPolicyTests
	{
		[Fact]
		public void PaymentHistoryRows_TreatRow13AsStartValueAndRow14AsFirstDataRow ()
		{
			Assert.Equal (12, AccountingPaymentHistoryPlanPolicy.HeaderRow);
			Assert.Equal (13, AccountingPaymentHistoryPlanPolicy.StartValueRow);
			Assert.Equal (14, AccountingPaymentHistoryPlanPolicy.FirstDataRow);
			Assert.False (AccountingPaymentHistoryPlanPolicy.IsDataRow (13));
			Assert.True (AccountingPaymentHistoryPlanPolicy.IsDataRow (14));
			Assert.True (AccountingPaymentHistoryPlanPolicy.IsDataRow (73));
			Assert.False (AccountingPaymentHistoryPlanPolicy.IsDataRow (74));
		}

		[Fact]
		public void EnsureWritableRow_RejectsHeaderStartAndOverflowRows ()
		{
			AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (14);
			AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (73);

			Assert.Throws<InvalidOperationException> (() => AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (12));
			Assert.Throws<InvalidOperationException> (() => AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (13));
			Assert.Throws<InvalidOperationException> (() => AccountingPaymentHistoryPlanPolicy.EnsureWritableRow (74));
		}

		[Fact]
		public void ResolveExpenseCharge_UsesPaymentBasisUntilExpenseBalanceIsExhausted ()
		{
			Assert.Equal (30000, AccountingPaymentHistoryPlanPolicy.ResolveExpenseCharge (30000, 120000));
			Assert.Equal (12000, AccountingPaymentHistoryPlanPolicy.ResolveExpenseCharge (30000, 12000));
			Assert.Equal (0, AccountingPaymentHistoryPlanPolicy.ResolveExpenseCharge (30000, 0));
			Assert.Equal (0, AccountingPaymentHistoryPlanPolicy.ResolveExpenseCharge (0, 12000));
		}

		[Fact]
		public void GetPreviousBalanceRowForDataRow_UsesImmediatelyPreviousRowWithoutTouchingHeader ()
		{
			Assert.Equal (13, AccountingPaymentHistoryPlanPolicy.GetPreviousBalanceRowForDataRow (14));
			Assert.Equal (14, AccountingPaymentHistoryPlanPolicy.GetPreviousBalanceRowForDataRow (15));
		}

		[Fact]
		public void IsDepositMarker_RecognizesOfficialAndLegacyAppliedText ()
		{
			Assert.True (AccountingPaymentHistoryPlanPolicy.IsDepositMarker ("（充当済み）"));
			Assert.True (AccountingPaymentHistoryPlanPolicy.IsDepositMarker (" （充当済み） "));
			Assert.True (AccountingPaymentHistoryPlanPolicy.IsDepositMarker ("充当済み"));
			Assert.True (AccountingPaymentHistoryPlanPolicy.IsDepositMarker ("(充当済み)"));
			Assert.False (AccountingPaymentHistoryPlanPolicy.IsDepositMarker ("済み"));
		}
	}
}
