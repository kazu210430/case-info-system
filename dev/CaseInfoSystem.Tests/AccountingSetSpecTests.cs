using CaseInfoSystem.ExcelAddIn.Domain;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class AccountingSetSpecTests
	{
		[Fact]
		public void LawyerReflectionTargets_IncludePaymentHistoryRowsA6ToA9()
		{
			Assert.Collection (
				AccountingSetSpec.LawyerReflectionTargets,
				target => AssertTarget (target, AccountingSetSpec.EstimateSheetName, AccountingSetSpec.LawyerWriteStartCellAddress),
				target => AssertTarget (target, AccountingSetSpec.InvoiceSheetName, AccountingSetSpec.LawyerWriteStartCellAddress),
				target => AssertTarget (target, AccountingSetSpec.ReceiptSheetName, AccountingSetSpec.LawyerWriteStartCellAddress),
				target => AssertTarget (target, AccountingSetSpec.AccountingRequestSheetName, AccountingSetSpec.LawyerWriteStartCellAddress),
				target => AssertTarget (target, AccountingSetSpec.PaymentHistorySheetName, AccountingSetSpec.PaymentHistoryLawyerWriteStartCellAddress));
		}

		private static void AssertTarget (AccountingLawyerReflectionTarget target, string sheetName, string startCellAddress)
		{
			Assert.Equal (sheetName, target.SheetName);
			Assert.Equal (startCellAddress, target.StartCellAddress);
		}
	}
}
