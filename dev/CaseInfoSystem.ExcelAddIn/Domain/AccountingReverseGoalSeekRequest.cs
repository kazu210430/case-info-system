namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class AccountingReverseGoalSeekRequest
    {
        internal AccountingReverseGoalSeekRequest(double targetAmount)
        {
            TargetAmount = targetAmount;
        }

        internal double TargetAmount { get; }
    }
}
