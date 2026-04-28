using System;
using System.Collections.Generic;
using System.Linq;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
	internal sealed class TemplateRegistrationValidationSummary
	{
		internal string TemplateDirectory { get; set; } = string.Empty;

		internal int DetectedFileCount { get; set; }

		internal List<TemplateRegistrationValidationEntry> TemplateResults { get; } = new List<TemplateRegistrationValidationEntry> ();

		internal int ValidTemplateCount => TemplateResults.Count (entry => entry != null && entry.IsValid);

		internal int ExcludedTemplateCount => TemplateResults.Count (entry => entry != null && !entry.IsValid);

		internal int WarningFileCount => TemplateResults.Count (entry => entry != null && entry.HasWarnings);

		internal IReadOnlyList<TemplateRegistrationValidationEntry> GetValidTemplates ()
		{
			return TemplateResults
				.Where (entry => entry != null && entry.IsValid)
				.OrderBy (entry => entry.Key ?? string.Empty, StringComparer.OrdinalIgnoreCase)
				.ThenBy (entry => entry.FileName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
				.ToList ();
		}
	}
}
