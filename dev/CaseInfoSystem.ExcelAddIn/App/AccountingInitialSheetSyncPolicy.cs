using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingInitialSheetSyncPolicy
	{
		internal const string WorkbookOpenEventName = ControlFlowReasons.WorkbookOpen;

		internal const string WorkbookActivateEventName = ControlFlowReasons.WorkbookActivate;

		internal const string WindowActivateEventName = ControlFlowReasons.WindowActivate;

		internal static bool ShouldSynchronizeActiveSheet (string eventName, bool isAccountingWorkbook)
		{
			return isAccountingWorkbook
				&& string.Equals (eventName, WindowActivateEventName, StringComparison.OrdinalIgnoreCase);
		}
	}
}
