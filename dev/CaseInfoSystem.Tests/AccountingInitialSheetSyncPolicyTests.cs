using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class AccountingInitialSheetSyncPolicyTests
	{
		[Theory]
		[InlineData("WindowActivate")]
		[InlineData("windowactivate")]
		public void ShouldSynchronizeActiveSheet_WhenAccountingWindowActivated_ReturnsTrue (string eventName)
		{
			Assert.True (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet (eventName, isAccountingWorkbook: true));
		}

		[Theory]
		[InlineData("WorkbookOpen")]
		[InlineData("WorkbookActivate")]
		[InlineData("Startup")]
		[InlineData("SheetActivate")]
		public void ShouldSynchronizeActiveSheet_WhenEventIsNotWindowActivate_ReturnsFalse (string eventName)
		{
			Assert.False (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet (eventName, isAccountingWorkbook: true));
		}

		[Fact]
		public void ShouldSynchronizeActiveSheet_WhenWorkbookIsNotAccounting_ReturnsFalse ()
		{
			Assert.False (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet ("WindowActivate", isAccountingWorkbook: false));
		}
	}
}
