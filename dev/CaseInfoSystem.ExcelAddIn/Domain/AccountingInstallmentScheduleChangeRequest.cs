namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class AccountingInstallmentScheduleChangeRequest
    {
        internal string BillingAmountText { get; set; }

        internal string ExpenseAmountText { get; set; }

        internal string WithholdingText { get; set; }

        internal string ChangeRoundText { get; set; }

        internal string ChangedInstallmentAmountText { get; set; }
    }
}
