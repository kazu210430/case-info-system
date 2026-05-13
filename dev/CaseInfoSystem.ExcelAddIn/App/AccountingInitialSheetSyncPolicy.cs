using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingInitialSheetSyncPolicy
	{
		internal const string WorkbookOpenEventName = "WorkbookOpen";

		internal const string WorkbookActivateEventName = "WorkbookActivate";

		internal static bool ShouldSynchronizeActiveSheet (string eventName, bool isAccountingWorkbook)
		{
			return isAccountingWorkbook
				&& string.Equals (eventName, WorkbookActivateEventName, StringComparison.OrdinalIgnoreCase);
		}
	}
}
