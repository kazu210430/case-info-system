using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class MergeDataBuilder
	{
		internal IReadOnlyDictionary<string, string> BuildMergeData (CaseContext caseContext)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			if (caseContext == null || caseContext.CaseValues == null) {
				return dictionary;
			}
			foreach (KeyValuePair<string, string> caseValue in caseContext.CaseValues) {
				string text = (caseValue.Key ?? string.Empty).Trim ();
				if (text.Length != 0) {
					dictionary [text] = NormalizeMergeValue (caseValue.Value ?? string.Empty);
				}
			}
			return dictionary;
		}

		private static string NormalizeMergeValue (string rawValue)
		{
			string value = ConvertFullWidthAsciiToHalfWidth (rawValue ?? string.Empty);
			return NormalizeForWordPlainText (value);
		}

		private static string ConvertFullWidthAsciiToHalfWidth (string value)
		{
			if (string.IsNullOrEmpty (value)) {
				return string.Empty;
			}
			char[] array = new char[value.Length];
			for (int i = 0; i < value.Length; i++) {
				char c = value [i];
				if (c == '\u3000') {
					array [i] = ' ';
				} else if (c >= '！' && c <= '～') {
					array [i] = (char)(c - 65248);
				} else {
					array [i] = c;
				}
			}
			return new string (array);
		}

		private static string NormalizeForWordPlainText (string value)
		{
			if (string.IsNullOrEmpty (value)) {
				return string.Empty;
			}
			return value.Replace ("\r\n", "\n").Replace ("\r", "\n").Replace ("\n", "\r");
		}
	}
}
