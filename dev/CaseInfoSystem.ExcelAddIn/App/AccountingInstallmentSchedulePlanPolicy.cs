using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingInstallmentSchedulePlanPolicy
	{
		internal const int HeaderRow = 12;
		internal const int StartValueRow = 13;
		internal const int FirstScheduleRow = 14;
		internal const int LastScheduleRow = 73;
		internal const int LastScheduleColumn = 10;
		internal const int PrintLastColumn = 9;

		internal static int MaxGeneratedRowCount {
			get { return LastScheduleRow - FirstScheduleRow + 1; }
		}

		internal static bool IsScheduleDetailRow (int rowNumber)
		{
			return rowNumber >= FirstScheduleRow && rowNumber <= LastScheduleRow;
		}

		internal static int GetPreviousBalanceRowForDetailRow (int detailRow)
		{
			EnsureWritableRow (detailRow);
			return detailRow - 1;
		}

		internal static DateTime AddOnePaymentMonth (DateTime value)
		{
			return value.AddMonths (1);
		}

		internal static double ResolveExpenseCharge (double paymentBasis, double previousExpenseBalance)
		{
			if (paymentBasis <= 0 || previousExpenseBalance <= 0) {
				return 0;
			}

			return Math.Min (paymentBasis, previousExpenseBalance);
		}

		internal static void EnsureWritableRow (int rowNumber)
		{
			if (rowNumber < FirstScheduleRow || rowNumber > LastScheduleRow) {
				throw new InvalidOperationException (
					"分割払い予定表の出力行がテーブル範囲 A12:J" + LastScheduleRow.ToString (System.Globalization.CultureInfo.InvariantCulture) +
					" を超えます。出力行: " + rowNumber.ToString (System.Globalization.CultureInfo.InvariantCulture));
			}
		}

		internal static AccountingInstallmentScheduleChangeStart ResolveChangeStart (
			IReadOnlyList<AccountingInstallmentScheduleExistingRow> rows,
			int changeRound)
		{
			if (changeRound < 1) {
				throw new InvalidOperationException ("変更回は 1 以上を入力してください。");
			}

			if (rows == null || rows.Count == 0) {
				throw new InvalidOperationException ("既存の分割払い予定表がありません。先に予定表を作成してください。");
			}

			AccountingInstallmentScheduleExistingRow target = null;
			foreach (AccountingInstallmentScheduleExistingRow row in rows) {
				if (row.Round >= changeRound) {
					target = row;
					break;
				}
			}

			if (target == null) {
				throw new InvalidOperationException ("変更回 " + changeRound.ToString (System.Globalization.CultureInfo.InvariantCulture) + " 以降の行が分割払い予定表に存在しません。");
			}

			if (target.RowNumber <= FirstScheduleRow) {
				throw new InvalidOperationException ("変更回が予定表の開始行です。途中変更として扱える既存行がありません。");
			}

			int previousRowNumber = target.RowNumber - 1;
			AccountingInstallmentScheduleExistingRow previous = FindByRowNumber (rows, previousRowNumber);
			if (previous == null) {
				throw new InvalidOperationException ("変更回の直前行 (" + previousRowNumber.ToString (System.Globalization.CultureInfo.InvariantCulture) + " 行目) が見つかりません。");
			}

			if (!IsUsableBalance (previous.BillingBalance)) {
				throw new InvalidOperationException ("変更回の直前行 (" + previousRowNumber.ToString (System.Globalization.CultureInfo.InvariantCulture) + " 行目) の請求残高が空欄または不正です。");
			}

			if (!IsUsableBalance (previous.ExpenseBalance) || previous.ExpenseBalance < 0) {
				throw new InvalidOperationException ("変更回の直前行 (" + previousRowNumber.ToString (System.Globalization.CultureInfo.InvariantCulture) + " 行目) の実費残高が不正です。");
			}

			if (previous.BillingBalance <= 0) {
				throw new InvalidOperationException ("変更回の直前ですでに請求残高が 0 円です。途中変更は反映できません。");
			}

			return new AccountingInstallmentScheduleChangeStart (target.RowNumber, previousRowNumber);
		}

		private static bool IsUsableBalance (double value)
		{
			return !double.IsNaN (value) && !double.IsInfinity (value);
		}

		private static AccountingInstallmentScheduleExistingRow FindByRowNumber (
			IReadOnlyList<AccountingInstallmentScheduleExistingRow> rows,
			int rowNumber)
		{
			foreach (AccountingInstallmentScheduleExistingRow row in rows) {
				if (row.RowNumber == rowNumber) {
					return row;
				}
			}

			return null;
		}
	}

	internal sealed class AccountingInstallmentScheduleExistingRow
	{
		internal AccountingInstallmentScheduleExistingRow (int rowNumber, int round, double billingBalance, double expenseBalance)
		{
			RowNumber = rowNumber;
			Round = round;
			BillingBalance = billingBalance;
			ExpenseBalance = expenseBalance;
		}

		internal int RowNumber { get; private set; }
		internal int Round { get; private set; }
		internal double BillingBalance { get; private set; }
		internal double ExpenseBalance { get; private set; }
	}

	internal sealed class AccountingInstallmentScheduleChangeStart
	{
		internal AccountingInstallmentScheduleChangeStart (int startRow, int previousRow)
		{
			StartRow = startRow;
			PreviousRow = previousRow;
		}

		internal int StartRow { get; private set; }
		internal int PreviousRow { get; private set; }
	}
}
