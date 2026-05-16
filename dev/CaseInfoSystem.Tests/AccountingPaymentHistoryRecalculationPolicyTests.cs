using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class AccountingPaymentHistoryRecalculationPolicyTests
	{
		[Fact]
		public void FirstRecalculationRow_StartsAfterDepositRow()
		{
			Assert.Equal (14, AccountingPaymentHistoryPlanPolicy.FirstDataRow);
			Assert.Equal (15, AccountingPaymentHistoryRecalculationPolicy.FirstRecalculationRow);
		}

		[Fact]
		public void ShouldRecalculateAfterSort_ReturnsFalse_WhenExpenseChargesDoNotIncrease()
		{
			Assert.False (AccountingPaymentHistoryRecalculationPolicy.ShouldRecalculateAfterSort (new double[0]));
			Assert.False (AccountingPaymentHistoryRecalculationPolicy.ShouldRecalculateAfterSort (new double[] { 1000 }));
			Assert.False (AccountingPaymentHistoryRecalculationPolicy.ShouldRecalculateAfterSort (new double[] { 1000, 1000, 500, 0 }));
		}

		[Fact]
		public void ShouldRecalculateAfterSort_ReturnsTrue_WhenLaterExpenseChargeIsGreater()
		{
			Assert.True (AccountingPaymentHistoryRecalculationPolicy.ShouldRecalculateAfterSort (new double[] { 1000, 1001 }));
			Assert.True (AccountingPaymentHistoryRecalculationPolicy.ShouldRecalculateAfterSort (new double[] { 1000, 500, 750 }));
		}

		[Fact]
		public void CommandService_CapturesCBeforeClearingH_AndUsesMemoizedTargetForGoalSeek()
		{
			string source = ReadAppSource ("AccountingPaymentHistoryCommandService.cs");

			int captureIndex = source.IndexOf ("Dictionary<int, double> targetReceiptAmounts = CaptureRecalculationTargetReceiptAmounts (workbook, lastDataRow);", StringComparison.Ordinal);
			int clearIndex = source.IndexOf ("ClearExpenseChargesForRecalculation (workbook, lastDataRow);", StringComparison.Ordinal);
			int targetIndex = source.IndexOf ("double targetReceiptAmount = targetReceiptAmounts[row];", StringComparison.Ordinal);
			int goalSeekIndex = source.IndexOf ("TryGoalSeekAndVerify (workbook, row, \"C\", \"D\", targetReceiptAmount, \"領収額\")", StringComparison.Ordinal);

			Assert.True (captureIndex >= 0, "C列目標値の退避が見つかりません。");
			Assert.True (clearIndex >= 0, "H列クリア処理が見つかりません。");
			Assert.True (targetIndex >= 0, "退避済みC列値の取得が見つかりません。");
			Assert.True (goalSeekIndex >= 0, "退避済みC列値をGoalSeek目標にする呼び出しが見つかりません。");
			Assert.True (captureIndex < clearIndex, "C列目標値はH列クリア前に退避してください。");
			Assert.True (targetIndex < goalSeekIndex, "GoalSeek目標値は退避済みC列値から取得してください。");
		}

		private static string ReadAppSource (string appFileName)
		{
			string repoRoot = FindRepositoryRoot ();
			return File.ReadAllText (Path.Combine (repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", appFileName));
		}

		private static string FindRepositoryRoot ()
		{
			DirectoryInfo current = new DirectoryInfo (Directory.GetCurrentDirectory ());
			while (current != null) {
				if (File.Exists (Path.Combine (current.FullName, "build.ps1"))
					&& Directory.Exists (Path.Combine (current.FullName, "dev", "CaseInfoSystem.ExcelAddIn"))) {
					return current.FullName;
				}

				current = current.Parent;
			}

			throw new DirectoryNotFoundException ("Repository root was not found.");
		}
	}
}
