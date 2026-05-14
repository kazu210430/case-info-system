using System;
using System.Globalization;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingPaymentHistoryPlanPolicy
	{
		internal const int HeaderRow = 12;
		internal const int StartValueRow = 13;
		internal const int FirstDataRow = 14;
		internal const int LastDataRow = 73;
		internal const int LastTableColumn = 10;
		internal const int PrintLastColumn = 9;
		internal const string DepositAppliedText = "（充当済み）";

		internal static bool IsDataRow (int rowNumber)
		{
			return rowNumber >= FirstDataRow && rowNumber <= LastDataRow;
		}

		internal static int GetPreviousBalanceRowForDataRow (int rowNumber)
		{
			EnsureWritableRow (rowNumber);
			return rowNumber - 1;
		}

		internal static double ResolveExpenseCharge (double paymentBasis, double previousExpenseBalance)
		{
			if (paymentBasis <= 0 || previousExpenseBalance <= 0) {
				return 0;
			}

			return Math.Min (paymentBasis, previousExpenseBalance);
		}

		internal static bool IsDepositMarker (string value)
		{
			string text = (value ?? string.Empty).Trim ();
			return string.Equals (text, DepositAppliedText, StringComparison.Ordinal);
		}

		internal static void EnsureWritableRow (int rowNumber)
		{
			if (!IsDataRow (rowNumber)) {
				throw new InvalidOperationException (
					"お支払い履歴の出力行がテーブル範囲 A12:J" + LastDataRow.ToString (CultureInfo.InvariantCulture) +
					" を超えます。出力行: " + rowNumber.ToString (CultureInfo.InvariantCulture));
			}
		}
	}
}
