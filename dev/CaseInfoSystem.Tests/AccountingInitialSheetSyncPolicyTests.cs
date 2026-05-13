using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class AccountingInitialSheetSyncPolicyTests
	{
		[Theory]
		[InlineData("WorkbookActivate")]
		[InlineData("workbookactivate")]
		public void ShouldSynchronizeActiveSheet_WhenAccountingWorkbookActivated_ReturnsTrue (string eventName)
		{
			Assert.True (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet (eventName, isAccountingWorkbook: true));
		}

		[Theory]
		[InlineData("WorkbookOpen")]
		[InlineData("Startup")]
		[InlineData("SheetActivate")]
		public void ShouldSynchronizeActiveSheet_WhenEventIsNotWorkbookActivate_ReturnsFalse (string eventName)
		{
			Assert.False (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet (eventName, isAccountingWorkbook: true));
		}

		[Fact]
		public void ShouldSynchronizeActiveSheet_WhenWorkbookIsNotAccounting_ReturnsFalse ()
		{
			Assert.False (AccountingInitialSheetSyncPolicy.ShouldSynchronizeActiveSheet ("WorkbookActivate", isAccountingWorkbook: false));
		}
	}
}
