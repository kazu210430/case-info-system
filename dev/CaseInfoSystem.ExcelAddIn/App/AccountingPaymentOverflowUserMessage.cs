using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingPaymentOverflowUserMessage
	{
		internal static readonly string UserMessage =
			"支払回数が60回を超えています。" +
			Environment.NewLine +
			"各回のお支払い額を増額してください。";

		private const string InstallmentScheduleOverflowDetail = "分割払い予定表が A12:J73 のテーブル範囲を超えます。分割金を増額してください。";
		private const string PaymentHistoryOutputOverflowPrefix = "お支払い履歴の出力行がテーブル範囲";
		private const string PaymentHistoryOutputRowLabel = "出力行:";

		internal static UserFacingException CreateInstallmentScheduleOverflowException (int attemptedRow, int lastRow, string tableRange)
		{
			string diagnostic =
				"AccountingInstallmentScheduleCommandService.CreateSchedule business validation failed." +
				Environment.NewLine +
				"procedure=AccountingInstallmentSchedule.CreateSchedule" +
				Environment.NewLine +
				"reason=InstallmentScheduleRowOverflow" +
				Environment.NewLine +
				"tableRange=" + (tableRange ?? string.Empty) +
				Environment.NewLine +
				"attemptedRow=" + attemptedRow.ToString (CultureInfo.InvariantCulture) +
				Environment.NewLine +
				"lastRow=" + lastRow.ToString (CultureInfo.InvariantCulture) +
				Environment.NewLine +
				"originalMessage=" + InstallmentScheduleOverflowDetail +
				Environment.NewLine +
				"sourceStackTrace=" + Environment.StackTrace;
			return new UserFacingException (UserMessage, diagnostic);
		}

		internal static bool IsPaymentHistoryOutputOverflow (Exception exception)
		{
			if (!(exception is InvalidOperationException) || exception is UserFacingException) {
				return false;
			}

			string message = exception.Message ?? string.Empty;
			return message.IndexOf (PaymentHistoryOutputOverflowPrefix, StringComparison.Ordinal) >= 0
				&& message.IndexOf (PaymentHistoryOutputRowLabel, StringComparison.Ordinal) >= 0;
		}

		internal static UserFacingException CreatePaymentHistoryOutputOverflowException (Exception exception)
		{
			if (exception == null) {
				throw new ArgumentNullException ("exception");
			}

			string diagnostic =
				"AccountingPaymentHistoryCommandService.OutputFutureBalance business validation failed." +
				Environment.NewLine +
				"procedure=AccountingPaymentHistory.OutputFutureBalance" +
				Environment.NewLine +
				"reason=PaymentHistoryOutputRowOverflow" +
				Environment.NewLine +
				"originalType=" + exception.GetType ().FullName +
				Environment.NewLine +
				"originalMessage=" + (exception.Message ?? string.Empty) +
				Environment.NewLine +
				"originalStackTrace=" + (exception.StackTrace ?? "(unavailable)");
			return new UserFacingException (UserMessage, diagnostic, exception);
		}
	}
}
