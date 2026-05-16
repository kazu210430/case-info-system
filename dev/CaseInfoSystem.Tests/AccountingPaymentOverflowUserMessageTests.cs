using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class AccountingPaymentOverflowUserMessageTests
    {
        private static readonly string ExpectedUserMessage =
            "支払回数が60回を超えています。" +
            Environment.NewLine +
            "各回のお支払い額を増額してください。";

        [Fact]
        public void InstallmentScheduleOverflowException_SeparatesUserMessageAndDiagnostic()
        {
            UserFacingException exception = AccountingPaymentOverflowUserMessage.CreateInstallmentScheduleOverflowException(74, 73, "A12:J73");

            Assert.Equal(ExpectedUserMessage, exception.UserMessage);
            Assert.DoesNotContain("エラーが発生しました", exception.UserMessage);
            Assert.DoesNotContain("AccountingInstallmentSchedule.CreateSchedule", exception.UserMessage);
            Assert.DoesNotContain("InvalidOperationException", exception.UserMessage);
            Assert.DoesNotContain("A12:J73", exception.UserMessage);
            Assert.Contains("procedure=AccountingInstallmentSchedule.CreateSchedule", exception.Message);
            Assert.Contains("reason=InstallmentScheduleRowOverflow", exception.Message);
            Assert.Contains("tableRange=A12:J73", exception.Message);
            Assert.Contains("attemptedRow=74", exception.Message);
            Assert.Contains("lastRow=73", exception.Message);
            Assert.Contains("originalMessage=分割払い予定表が A12:J73 のテーブル範囲を超えます。分割金を増額してください。", exception.Message);
            Assert.Contains("sourceStackTrace=", exception.Message);
            Assert.Null(exception.InnerException);
        }

        [Fact]
        public void PaymentHistoryOutputOverflowException_SeparatesUserMessageAndKeepsOriginalException()
        {
            var original = new InvalidOperationException("お支払い履歴の出力行がテーブル範囲 A12:J73 を超えます。出力行: 74");

            UserFacingException exception = AccountingPaymentOverflowUserMessage.CreatePaymentHistoryOutputOverflowException(original);

            Assert.Equal(ExpectedUserMessage, exception.UserMessage);
            Assert.DoesNotContain("エラーが発生しました", exception.UserMessage);
            Assert.DoesNotContain("AccountingPaymentHistory.OutputFutureBalance", exception.UserMessage);
            Assert.DoesNotContain("InvalidOperationException", exception.UserMessage);
            Assert.DoesNotContain("A12:J73", exception.UserMessage);
            Assert.DoesNotContain("出力行", exception.UserMessage);
            Assert.Same(original, exception.InnerException);
            Assert.Contains("procedure=AccountingPaymentHistory.OutputFutureBalance", exception.Message);
            Assert.Contains("reason=PaymentHistoryOutputRowOverflow", exception.Message);
            Assert.Contains("originalType=System.InvalidOperationException", exception.Message);
            Assert.Contains("originalMessage=" + original.Message, exception.Message);
            Assert.Contains("originalStackTrace=", exception.Message);
        }

        [Fact]
        public void PaymentHistoryOutputOverflowDetection_OnlyMatchesTheKnownRangeOverflow()
        {
            Assert.True(AccountingPaymentOverflowUserMessage.IsPaymentHistoryOutputOverflow(
                new InvalidOperationException("お支払い履歴の出力行がテーブル範囲 A12:J73 を超えます。出力行: 74")));
            Assert.False(AccountingPaymentOverflowUserMessage.IsPaymentHistoryOutputOverflow(
                new InvalidOperationException("別の入力エラーです。")));
            Assert.False(AccountingPaymentOverflowUserMessage.IsPaymentHistoryOutputOverflow(
                new UserFacingException(ExpectedUserMessage, "diagnostic")));
        }

        [Fact]
        public void CommandServices_RouteOverflowCasesThroughSharedUserFacingMessage()
        {
            string installmentSource = ReadAppSource("AccountingInstallmentScheduleCommandService.cs");
            string paymentHistorySource = ReadAppSource("AccountingPaymentHistoryCommandService.cs");

            Assert.Contains("AccountingPaymentOverflowUserMessage.CreateInstallmentScheduleOverflowException", installmentSource);
            Assert.DoesNotContain("throw new InvalidOperationException (\"分割払い予定表が A12:J73 のテーブル範囲を超えます。分割金を増額してください。\")", installmentSource);
            Assert.Contains("AccountingPaymentOverflowUserMessage.IsPaymentHistoryOutputOverflow (exception)", paymentHistorySource);
            Assert.Contains("AccountingPaymentOverflowUserMessage.CreatePaymentHistoryOutputOverflowException (exception)", paymentHistorySource);
            Assert.Contains("_userErrorService.ShowUserError (\"AccountingPaymentHistory.OutputFutureBalance\", userErrorException)", paymentHistorySource);
        }

        [Fact]
        public void MissingPaymentHistoryImportMessage_RemainsUserFacingException()
        {
            string source = ReadAppSource("AccountingPaymentHistoryImportService.cs");

            Assert.Contains("private const string PaymentHistoryRequiredMessage = \"お支払い履歴を先に作成してください。\";", source);
            Assert.Contains("new UserFacingException (PaymentHistoryRequiredMessage, diagnostic)", source);
            Assert.DoesNotContain("throw new InvalidOperationException (\"お支払い履歴を先に作成してください。\")", source);
        }

        private static string ReadAppSource(string appFileName)
        {
            string repoRoot = FindRepositoryRoot();
            return File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", appFileName));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
