namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingPaymentHistoryRecalculationPolicy
	{
		internal const int FirstRecalculationRow = 15;

		internal static bool IsInsertedRowInRecalculationRange (int appendedRow, int sortedInsertedRow)
		{
			return appendedRow >= FirstRecalculationRow
				&& sortedInsertedRow >= FirstRecalculationRow
				&& sortedInsertedRow < appendedRow;
		}

		internal static bool ShouldRecalculateInsertedRow (int appendedRow, int sortedInsertedRow, double previousExpenseBalance)
		{
			return IsInsertedRowInRecalculationRange (appendedRow, sortedInsertedRow)
				&& previousExpenseBalance > 0;
		}
	}
}
