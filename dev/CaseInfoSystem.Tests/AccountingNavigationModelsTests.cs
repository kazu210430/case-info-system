using System.Linq;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class AccountingNavigationModelsTests
    {
        private const string SectionAction = "実行";
        private const string SectionNavigation = "画面切替";

        [Theory]
        [InlineData("見積書", "set-issue-date", 2, "見積書を取り込む（表示中）")]
        [InlineData("請求書", "set-issue-date-and-due-date", 3, "請求書を取り込む（表示中）")]
        [InlineData("領収書", "set-issue-date", 4, "領収書を取り込む（表示中）")]
        public void CreateForSheet_WhenAccountingDocumentSheet_BuildsExecutionActionsInExpectedOrder(
            string activeSheetCodeName,
            string expectedIssueDateActionId,
            int disabledImportIndex,
            string expectedCurrentImportCaption)
        {
            var actions = AccountingNavigationDefinitions.CreateForSheet(activeSheetCodeName);
            var executionActions = actions.Where(action => action.SectionTitle == SectionAction).ToArray();

            string[] expectedActionIds =
            {
                expectedIssueDateActionId,
                "show-save-as-prompt",
                "import-estimate-to-current",
                "import-invoice-to-current",
                "import-receipt-to-current",
                "open-reverse-tool",
                "reset-current-sheet"
            };
            string[] expectedCaptions =
            {
                "発行日を入力",
                "別名保存",
                "見積書を取り込む",
                "請求書を取り込む",
                "領収書を取り込む",
                "逆算ツール",
                "リセット"
            };
            bool[] expectedEnabled =
            {
                true,
                true,
                true,
                true,
                true,
                true,
                true
            };
            expectedCaptions[disabledImportIndex] = expectedCurrentImportCaption;
            expectedEnabled[disabledImportIndex] = false;

            Assert.Equal(expectedActionIds, executionActions.Select(action => action.ActionId).ToArray());
            Assert.Equal(expectedCaptions, executionActions.Select(action => action.Caption).ToArray());
            Assert.Equal(expectedEnabled, executionActions.Select(action => action.IsEnabled).ToArray());
        }

        [Theory]
        [InlineData("見積書", 0, "見積書（表示中）")]
        [InlineData("請求書", 1, "請求書（表示中）")]
        [InlineData("領収書", 2, "領収書（表示中）")]
        public void CreateForSheet_WhenAccountingDocumentSheet_KeepsScreenSwitchActions(
            string activeSheetCodeName,
            int currentNavigationIndex,
            string expectedCurrentCaption)
        {
            var actions = AccountingNavigationDefinitions.CreateForSheet(activeSheetCodeName);
            var navigationActions = actions.Where(action => action.SectionTitle == SectionNavigation).ToArray();

            string[] expectedActionIds =
            {
                "switch-to-estimate-sheet",
                "switch-to-invoice-sheet",
                "switch-to-receipt-sheet",
                "switch-to-accounting-request-sheet",
                "switch-to-installment-sheet",
                "switch-to-payment-history-sheet"
            };
            string[] expectedCaptions =
            {
                "見積書",
                "請求書",
                "領収書",
                "会計依頼書",
                "分割払い予定表",
                "お支払い履歴"
            };
            bool[] expectedEnabled =
            {
                true,
                true,
                true,
                true,
                true,
                true
            };
            expectedCaptions[currentNavigationIndex] = expectedCurrentCaption;
            expectedEnabled[currentNavigationIndex] = false;

            Assert.Equal(expectedActionIds, navigationActions.Select(action => action.ActionId).ToArray());
            Assert.Equal(expectedCaptions, navigationActions.Select(action => action.Caption).ToArray());
            Assert.Equal(expectedEnabled, navigationActions.Select(action => action.IsEnabled).ToArray());
        }

        [Fact]
        public void CreateForSheet_WhenAccountingRequestSheet_BuildsExecutionActionsInRequestedOrder()
        {
            var actions = AccountingNavigationDefinitions.CreateForSheet("会計依頼書");
            var executionActions = actions.Where(action => action.SectionTitle == SectionAction).ToArray();

            string[] expectedActionIds =
            {
                "set-issue-date",
                "show-save-as-prompt",
                "import-estimate-to-current",
                "import-invoice-to-current",
                "import-receipt-to-current",
                "import-payment-history-to-request",
                "open-reverse-tool",
                "reset-current-sheet"
            };
            string[] expectedCaptions =
            {
                "発行日を入力",
                "別名保存",
                "見積書を取り込む",
                "請求書を取り込む",
                "領収書を取り込む",
                "お支払い履歴を取り込む",
                "逆算ツール",
                "リセット"
            };

            Assert.Equal(expectedActionIds, executionActions.Select(action => action.ActionId).ToArray());
            Assert.Equal(expectedCaptions, executionActions.Select(action => action.Caption).ToArray());
            Assert.All(executionActions, action => Assert.True(action.IsEnabled));
        }

        [Fact]
        public void CreateForSheet_WhenAccountingRequestSheet_KeepsScreenSwitchActions()
        {
            var actions = AccountingNavigationDefinitions.CreateForSheet("会計依頼書");
            var navigationActions = actions.Where(action => action.SectionTitle == SectionNavigation).ToArray();

            string[] expectedActionIds =
            {
                "switch-to-estimate-sheet",
                "switch-to-invoice-sheet",
                "switch-to-receipt-sheet",
                "switch-to-accounting-request-sheet",
                "switch-to-installment-sheet",
                "switch-to-payment-history-sheet"
            };
            string[] expectedCaptions =
            {
                "見積書",
                "請求書",
                "領収書",
                "会計依頼書（表示中）",
                "分割払い予定表",
                "お支払い履歴"
            };
            bool[] expectedEnabled =
            {
                true,
                true,
                true,
                false,
                true,
                true
            };

            Assert.Equal(expectedActionIds, navigationActions.Select(action => action.ActionId).ToArray());
            Assert.Equal(expectedCaptions, navigationActions.Select(action => action.Caption).ToArray());
            Assert.Equal(expectedEnabled, navigationActions.Select(action => action.IsEnabled).ToArray());
        }

        [Fact]
        public void CreateForSheet_WhenPaymentHistorySheet_BuildsExecutionActionsBeforeScreenSwitch()
        {
            var actions = AccountingNavigationDefinitions.CreateForSheet("お支払い履歴");
            var sectionOrder = actions.Select(action => action.SectionTitle).Distinct().ToArray();
            var executionActions = actions.Where(action => action.SectionTitle == SectionAction).ToArray();
            var navigationActions = actions.Where(action => action.SectionTitle == SectionNavigation).ToArray();

            string[] expectedExecutionActionIds =
            {
                "set-payment-history-issue-date",
                "show-save-as-prompt",
                "reset-payment-history"
            };
            string[] expectedExecutionCaptions =
            {
                "発行日を入力",
                "別名保存",
                "リセット"
            };
            string[] expectedNavigationActionIds =
            {
                "switch-to-estimate-sheet",
                "switch-to-invoice-sheet",
                "switch-to-receipt-sheet",
                "switch-to-accounting-request-sheet",
                "switch-to-installment-sheet",
                "switch-to-payment-history-sheet"
            };
            string[] expectedNavigationCaptions =
            {
                "見積書",
                "請求書",
                "領収書",
                "会計依頼書",
                "分割払い予定表",
                "お支払い履歴（表示中）"
            };
            bool[] expectedNavigationEnabled =
            {
                true,
                true,
                true,
                true,
                true,
                false
            };

            Assert.Equal(new[] { SectionAction, SectionNavigation }, sectionOrder);
            Assert.Equal(expectedExecutionActionIds, executionActions.Select(action => action.ActionId).ToArray());
            Assert.Equal(expectedExecutionCaptions, executionActions.Select(action => action.Caption).ToArray());
            Assert.All(executionActions, action => Assert.True(action.IsEnabled));
            Assert.Equal(expectedNavigationActionIds, navigationActions.Select(action => action.ActionId).ToArray());
            Assert.Equal(expectedNavigationCaptions, navigationActions.Select(action => action.Caption).ToArray());
            Assert.Equal(expectedNavigationEnabled, navigationActions.Select(action => action.IsEnabled).ToArray());
        }
    }
}
