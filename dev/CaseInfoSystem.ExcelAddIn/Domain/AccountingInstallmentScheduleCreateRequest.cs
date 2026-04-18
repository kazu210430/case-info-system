namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class AccountingInstallmentScheduleCreateRequest
    {
        internal string BillingAmountText { get; set; }

        internal string ExpenseAmountText { get; set; }

        internal string WithholdingText { get; set; }

        internal string FirstDueDateText { get; set; }

        internal string DepositAmountText { get; set; }

        internal string InstallmentAmountText { get; set; }
    }
}
