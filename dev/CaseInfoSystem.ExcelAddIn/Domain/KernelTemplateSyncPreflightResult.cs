namespace CaseInfoSystem.ExcelAddIn.Domain
{
	internal sealed class KernelTemplateSyncPreflightResult
	{
		private KernelTemplateSyncPreflightResult (KernelTemplateSyncPreflightStatus status, string templateDirectory, TemplateRegistrationValidationSummary validationSummary, ValidationFailureSummary failure)
		{
			Status = status;
			TemplateDirectory = templateDirectory ?? string.Empty;
			ValidationSummary = validationSummary;
			Failure = failure;
		}

		internal KernelTemplateSyncPreflightStatus Status { get; }

		internal string TemplateDirectory { get; }

		internal TemplateRegistrationValidationSummary ValidationSummary { get; }

		internal ValidationFailureSummary Failure { get; }

		internal static KernelTemplateSyncPreflightResult Succeeded (string templateDirectory, TemplateRegistrationValidationSummary validationSummary)
		{
			return new KernelTemplateSyncPreflightResult (KernelTemplateSyncPreflightStatus.Succeeded, templateDirectory, validationSummary, null);
		}

		internal static KernelTemplateSyncPreflightResult Failed (string templateDirectory, ValidationFailureSummary failure, TemplateRegistrationValidationSummary validationSummary = null)
		{
			return new KernelTemplateSyncPreflightResult (KernelTemplateSyncPreflightStatus.Failed, templateDirectory, validationSummary, failure);
		}
	}
}
