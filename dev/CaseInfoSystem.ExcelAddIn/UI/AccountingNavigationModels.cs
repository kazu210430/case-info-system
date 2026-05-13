using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    // クラス: 会計ナビゲーション上の1アクション定義を表す。
    // 責務: ボタン識別子、表示名、所属セクション、有効状態を保持する。
    internal sealed class AccountingNavigationActionDefinition
    {
        // メソッド: 会計ナビゲーションアクション定義を初期化する。
        // 引数: actionId - 動作識別子, caption - 表示名, sectionTitle - セクション名, isEnabled - 有効状態。
        // 戻り値: なし。
        // 副作用: 内部状態を初期化する。
        internal AccountingNavigationActionDefinition(string actionId, string caption, string sectionTitle, bool isEnabled)
        {
            ActionId = actionId ?? string.Empty;
            Caption = caption ?? string.Empty;
            SectionTitle = sectionTitle ?? string.Empty;
            IsEnabled = isEnabled;
        }

        internal string ActionId { get; }

        internal string Caption { get; }

        internal string SectionTitle { get; }

        internal bool IsEnabled { get; }
    }

    // クラス: 会計ナビゲーションのアクションID定数をまとめる。
    // 責務: 画面実行・画面切替で利用する識別子を集中管理する。
    internal static class AccountingNavigationActionIds
    {
        internal const string ShowSaveAsPrompt = "show-save-as-prompt";
        internal const string ImportPaymentHistoryToRequest = "import-payment-history-to-request";
        internal const string ImportEstimateToCurrent = "import-estimate-to-current";
        internal const string ImportInvoiceToCurrent = "import-invoice-to-current";
        internal const string ImportReceiptToCurrent = "import-receipt-to-current";
        internal const string OpenPaymentHistoryInput = "open-payment-history-input";
        internal const string OpenInstallmentScheduleInput = "open-installment-schedule-input";
        internal const string ResetCurrentSheet = "reset-current-sheet";
        internal const string SetIssueDate = "set-issue-date";
        internal const string SetIssueDateAndDueDate = "set-issue-date-and-due-date";
        internal const string OpenReverseTool = "open-reverse-tool";
        internal const string SwitchToEstimateSheet = "switch-to-estimate-sheet";
        internal const string SwitchToInvoiceSheet = "switch-to-invoice-sheet";
        internal const string SwitchToReceiptSheet = "switch-to-receipt-sheet";
        internal const string SwitchToAccountingRequestSheet = "switch-to-accounting-request-sheet";
        internal const string SwitchToInstallmentSheet = "switch-to-installment-sheet";
        internal const string SwitchToPaymentHistorySheet = "switch-to-payment-history-sheet";
    }

    // クラス: アクティブシートに応じた会計ナビゲーション定義を組み立てる。
    // 責務: 実行アクションとシート切替アクションを画面状態に合わせて返す。
    internal static class AccountingNavigationDefinitions
    {
        private const string SectionAction = "実行";
        private const string SectionNavigation = "画面切替";

        // メソッド: アクティブシートに応じたナビゲーション定義一覧を返す。
        // 引数: activeSheetCodeName - 現在シートの CodeName。
        // 戻り値: 表示対象のアクション定義一覧。
        // 副作用: なし。
        internal static IReadOnlyList<AccountingNavigationActionDefinition> CreateForSheet(string activeSheetCodeName)
        {
            var definitions = new List<AccountingNavigationActionDefinition>();

            if (string.Equals(activeSheetCodeName, Domain.AccountingSetSpec.EstimateSheetName, StringComparison.OrdinalIgnoreCase))
            {
                AddAccountingDocumentExecutionActions(definitions, AccountingNavigationActionIds.SetIssueDate, Domain.AccountingSetSpec.EstimateSheetName);
                AddSheetSwitchActions(definitions, Domain.AccountingSetSpec.EstimateSheetName);
                return definitions;
            }

            if (string.Equals(activeSheetCodeName, Domain.AccountingSetSpec.InvoiceSheetName, StringComparison.OrdinalIgnoreCase))
            {
                AddAccountingDocumentExecutionActions(definitions, AccountingNavigationActionIds.SetIssueDateAndDueDate, Domain.AccountingSetSpec.InvoiceSheetName);
                AddSheetSwitchActions(definitions, Domain.AccountingSetSpec.InvoiceSheetName);
                return definitions;
            }

            if (string.Equals(activeSheetCodeName, Domain.AccountingSetSpec.ReceiptSheetName, StringComparison.OrdinalIgnoreCase))
            {
                AddAccountingDocumentExecutionActions(definitions, AccountingNavigationActionIds.SetIssueDate, Domain.AccountingSetSpec.ReceiptSheetName);
                AddSheetSwitchActions(definitions, Domain.AccountingSetSpec.ReceiptSheetName);
                return definitions;
            }

            if (string.Equals(activeSheetCodeName, Domain.AccountingSetSpec.AccountingRequestSheetName, StringComparison.OrdinalIgnoreCase))
            {
                AddMainFormExecutionActions(definitions, AccountingNavigationActionIds.SetIssueDate, true, true, true, true);
                AddSheetSwitchActions(definitions, Domain.AccountingSetSpec.AccountingRequestSheetName);
                return definitions;
            }

            if (string.Equals(activeSheetCodeName, Domain.AccountingSetSpec.InstallmentSheetName, StringComparison.OrdinalIgnoreCase))
            {
                AddSheetSwitchActions(definitions, Domain.AccountingSetSpec.InstallmentSheetName);
                return definitions;
            }

            if (string.Equals(activeSheetCodeName, Domain.AccountingSetSpec.PaymentHistorySheetName, StringComparison.OrdinalIgnoreCase))
            {
                AddSheetSwitchActions(definitions, Domain.AccountingSetSpec.PaymentHistorySheetName);
                return definitions;
            }

            return definitions;
        }

        // メソッド: 見積書・請求書・領収書向けの実行アクションを追加する。
        // 引数: definitions - 追加先, issueDateActionId - 発行日系動作ID, activeSheetCodeName - 現在シートの CodeName。
        // 戻り値: なし。
        // 副作用: definitions を更新する。
        private static void AddAccountingDocumentExecutionActions(
            ICollection<AccountingNavigationActionDefinition> definitions,
            string issueDateActionId,
            string activeSheetCodeName)
        {
            definitions.Add(new AccountingNavigationActionDefinition(
                string.IsNullOrWhiteSpace(issueDateActionId) ? AccountingNavigationActionIds.SetIssueDate : issueDateActionId,
                "発行日を入力",
                SectionAction,
                true));
            definitions.Add(new AccountingNavigationActionDefinition(
                AccountingNavigationActionIds.ShowSaveAsPrompt,
                "別名保存",
                SectionAction,
                true));

            AddImportAction(definitions, AccountingNavigationActionIds.ImportEstimateToCurrent, "見積書を取り込む", Domain.AccountingSetSpec.EstimateSheetName, activeSheetCodeName);
            AddImportAction(definitions, AccountingNavigationActionIds.ImportInvoiceToCurrent, "請求書を取り込む", Domain.AccountingSetSpec.InvoiceSheetName, activeSheetCodeName);
            AddImportAction(definitions, AccountingNavigationActionIds.ImportReceiptToCurrent, "領収書を取り込む", Domain.AccountingSetSpec.ReceiptSheetName, activeSheetCodeName);

            definitions.Add(new AccountingNavigationActionDefinition(
                AccountingNavigationActionIds.OpenReverseTool,
                "逆算ツール",
                SectionAction,
                true));
            definitions.Add(new AccountingNavigationActionDefinition(
                AccountingNavigationActionIds.ResetCurrentSheet,
                "リセット",
                SectionAction,
                true));
        }

        // メソッド: メイン帳票系シート向けの実行アクションを追加する。
        // 引数: definitions - 追加先, issueDateActionId - 発行日系動作ID, includeEstimateImport 等 - 各取込ボタンを含めるか。
        // 戻り値: なし。
        // 副作用: definitions を更新する。
        private static void AddMainFormExecutionActions(
            ICollection<AccountingNavigationActionDefinition> definitions,
            string issueDateActionId,
            bool includeEstimateImport,
            bool includeInvoiceImport,
            bool includeReceiptImport,
            bool includePaymentHistoryImport)
        {
            definitions.Add(new AccountingNavigationActionDefinition(
                string.IsNullOrWhiteSpace(issueDateActionId) ? AccountingNavigationActionIds.SetIssueDate : issueDateActionId,
                "発行日を入力",
                SectionAction,
                true));
            definitions.Add(new AccountingNavigationActionDefinition(
                AccountingNavigationActionIds.OpenReverseTool,
                "逆算ツール",
                SectionAction,
                true));
            definitions.Add(new AccountingNavigationActionDefinition(
                AccountingNavigationActionIds.ShowSaveAsPrompt,
                "別名保存",
                SectionAction,
                true));

            if (includeEstimateImport)
            {
                definitions.Add(new AccountingNavigationActionDefinition(
                    AccountingNavigationActionIds.ImportEstimateToCurrent,
                    "見積書を取り込む",
                    SectionAction,
                    true));
            }

            if (includeInvoiceImport)
            {
                definitions.Add(new AccountingNavigationActionDefinition(
                    AccountingNavigationActionIds.ImportInvoiceToCurrent,
                    "請求書を取り込む",
                    SectionAction,
                    true));
            }

            if (includeReceiptImport)
            {
                definitions.Add(new AccountingNavigationActionDefinition(
                    AccountingNavigationActionIds.ImportReceiptToCurrent,
                    "領収書を取り込む",
                    SectionAction,
                    true));
            }

            if (includePaymentHistoryImport)
            {
                definitions.Add(new AccountingNavigationActionDefinition(
                    AccountingNavigationActionIds.ImportPaymentHistoryToRequest,
                    "お支払い履歴を取り込む",
                    SectionAction,
                    true));
            }

            definitions.Add(new AccountingNavigationActionDefinition(
                AccountingNavigationActionIds.ResetCurrentSheet,
                "リセット",
                SectionAction,
                true));
        }

        private static void AddImportAction(
            ICollection<AccountingNavigationActionDefinition> definitions,
            string actionId,
            string caption,
            string targetSheetCodeName,
            string activeSheetCodeName)
        {
            bool isCurrent = string.Equals(targetSheetCodeName, activeSheetCodeName, StringComparison.OrdinalIgnoreCase);
            definitions.Add(new AccountingNavigationActionDefinition(
                actionId,
                isCurrent ? caption + "（表示中）" : caption,
                SectionAction,
                !isCurrent));
        }

        // メソッド: シート切替アクション群を追加する。
        // 引数: definitions - 追加先, activeSheetCodeName - 現在シートの CodeName。
        // 戻り値: なし。
        // 副作用: definitions を更新する。
        private static void AddSheetSwitchActions(ICollection<AccountingNavigationActionDefinition> definitions, string activeSheetCodeName)
        {
            AddSheetSwitchAction(definitions, AccountingNavigationActionIds.SwitchToEstimateSheet, "見積書", Domain.AccountingSetSpec.EstimateSheetName, activeSheetCodeName);
            AddSheetSwitchAction(definitions, AccountingNavigationActionIds.SwitchToInvoiceSheet, "請求書", Domain.AccountingSetSpec.InvoiceSheetName, activeSheetCodeName);
            AddSheetSwitchAction(definitions, AccountingNavigationActionIds.SwitchToReceiptSheet, "領収書", Domain.AccountingSetSpec.ReceiptSheetName, activeSheetCodeName);
            AddSheetSwitchAction(definitions, AccountingNavigationActionIds.SwitchToAccountingRequestSheet, "会計依頼書", Domain.AccountingSetSpec.AccountingRequestSheetName, activeSheetCodeName);
            AddSheetSwitchAction(definitions, AccountingNavigationActionIds.SwitchToInstallmentSheet, "分割払い予定表", Domain.AccountingSetSpec.InstallmentSheetName, activeSheetCodeName);
            AddSheetSwitchAction(definitions, AccountingNavigationActionIds.SwitchToPaymentHistorySheet, "お支払い履歴", Domain.AccountingSetSpec.PaymentHistorySheetName, activeSheetCodeName);
        }

        private static void AddSheetSwitchAction(
            ICollection<AccountingNavigationActionDefinition> definitions,
            string actionId,
            string caption,
            string targetSheetCodeName,
            string activeSheetCodeName)
        {
            bool isCurrent = string.Equals(targetSheetCodeName, activeSheetCodeName, StringComparison.OrdinalIgnoreCase);
            definitions.Add(new AccountingNavigationActionDefinition(
                actionId,
                isCurrent ? caption + "（表示中）" : caption,
                SectionNavigation,
                !isCurrent));
        }
    }

    // クラス: 会計ナビゲーションで押下されたアクションIDを通知するイベント引数。
    // 責務: 選択されたアクションIDを保持する。
    internal sealed class AccountingNavigationActionEventArgs : EventArgs
    {
        // メソッド: イベント引数を初期化する。
        // 引数: actionId - 選択されたアクションID。
        // 戻り値: なし。
        // 副作用: 内部状態を初期化する。
        internal AccountingNavigationActionEventArgs(string actionId)
        {
            ActionId = actionId ?? string.Empty;
        }

        internal string ActionId { get; }
    }
}

