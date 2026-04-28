using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
	internal sealed class TemplateRegistrationValidationEntry
	{
		private readonly List<string> _errors = new List<string> ();

		private readonly List<string> _warnings = new List<string> ();

		internal string FileName { get; set; } = string.Empty;

		internal string Key { get; set; } = string.Empty;

		internal string DisplayName { get; set; } = string.Empty;

		internal IReadOnlyList<string> Errors => _errors;

		internal IReadOnlyList<string> Warnings => _warnings;

		internal bool IsValid => _errors.Count == 0;

		internal bool HasWarnings => _warnings.Count > 0;

		internal void AddError (string message)
		{
			if (string.IsNullOrWhiteSpace (message)) {
				return;
			}
			if (!_errors.Contains (message)) {
				_errors.Add (message);
			}
		}

		internal void AddWarning (string message)
		{
			if (string.IsNullOrWhiteSpace (message)) {
				return;
			}
			if (!_warnings.Contains (message)) {
				_warnings.Add (message);
			}
		}
	}
}
