namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class AccountingGoalSeekExecutionResult
	{
		internal AccountingGoalSeekExecutionResult (bool succeeded, object currentValue)
		{
			Succeeded = succeeded;
			CurrentValue = currentValue;
		}

		internal bool Succeeded { get; private set; }

		internal object CurrentValue { get; private set; }
	}
}
