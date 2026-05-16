using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingPaymentHistoryRecalculationPolicy
	{
		internal const int FirstRecalculationRow = 15;

		internal static bool ShouldRecalculateAfterSort (IList<double> expenseCharges)
		{
			if (expenseCharges == null) {
				throw new ArgumentNullException ("expenseCharges");
			}

			for (int index = 1; index < expenseCharges.Count; index++) {
				if (expenseCharges[index - 1] < expenseCharges[index]) {
					return true;
				}
			}

			return false;
		}
	}
}
