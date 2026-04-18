namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class AccountingPaymentHistoryEntryRequest
    {
        internal string BillingAmountText { get; set; }

        internal string ExpenseAmountText { get; set; }

        internal string WithholdingText { get; set; }

        internal string DepositAmountText { get; set; }

        internal string ReceiptDateText { get; set; }

        internal string ReceiptAmountText { get; set; }
    }
}
