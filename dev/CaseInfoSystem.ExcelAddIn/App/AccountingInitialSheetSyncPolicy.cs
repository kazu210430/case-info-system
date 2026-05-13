using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingInitialSheetSyncPolicy
	{
		internal const string WorkbookOpenEventName = "WorkbookOpen";

		internal const string WorkbookActivateEventName = "WorkbookActivate";

		internal const string WindowActivateEventName = "WindowActivate";

		internal static bool ShouldSynchronizeActiveSheet (string eventName, bool isAccountingWorkbook)
		{
			return isAccountingWorkbook
				&& string.Equals (eventName, WindowActivateEventName, StringComparison.OrdinalIgnoreCase);
		}
	}
}
