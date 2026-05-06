using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
	internal sealed class ValidationFailureSummary
	{
		internal ValidationFailureSummary (ValidationFailureKind kind, string message, int detectedCount, IReadOnlyList<TemplateRegistrationValidationEntry> templateResults)
		{
			Kind = kind;
			Message = message ?? string.Empty;
			DetectedCount = detectedCount;
			TemplateResults = templateResults ?? Array.Empty<TemplateRegistrationValidationEntry> ();
		}

		internal ValidationFailureKind Kind { get; }

		internal string Message { get; }

		internal int DetectedCount { get; }

		internal IReadOnlyList<TemplateRegistrationValidationEntry> TemplateResults { get; }
	}
}
