using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class AccountingIssueDateDueDatePolicyTests
	{
		[Theory]
		[InlineData (0.0, true)]
		[InlineData (1.0, false)]
		[InlineData (-1.0, false)]
		public void ShouldWriteNoPaymentNotice_ReturnsTrueOnlyForZero (double invoicePaymentAmount, bool expected)
		{
			bool result = AccountingIssueDateDueDatePolicy.ShouldWriteNoPaymentNotice (invoicePaymentAmount);

			Assert.Equal (expected, result);
		}

		[Fact]
		public void NoPaymentNoticeText_MatchesSpecifiedMessage ()
		{
			Assert.Equal ("（お支払い頂く金額はございません）", AccountingIssueDateDueDatePolicy.NoPaymentNoticeText);
		}
	}
}
