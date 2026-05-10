using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class FieldKeyRenameMap
	{
		internal const string LegacyLawyerKey = "当方_弁護士";
		internal const string CurrentLawyerKey = "My表示名";
		internal const string LegacyPostalCodeKey = "当方_郵便番号";
		internal const string CurrentPostalCodeKey = "My郵便番号";
		internal const string LegacyAddressKey = "当方_住所";
		internal const string CurrentAddressKey = "My住所";
		internal const string LegacyOfficeNameKey = "当方_事務所名";
		internal const string CurrentOfficeNameKey = "My事務所名";
		internal const string LegacyPhoneKey = "当方_電話";
		internal const string CurrentPhoneKey = "My電話";
		internal const string LegacyFaxKey = "当方_Fax";
		internal const string CurrentFaxKey = "MyFax";

		internal static readonly string[] LegacyAndCurrentFieldKeys =
		{
			LegacyLawyerKey,
			CurrentLawyerKey,
			LegacyPostalCodeKey,
			CurrentPostalCodeKey,
			LegacyAddressKey,
			CurrentAddressKey,
			LegacyOfficeNameKey,
			CurrentOfficeNameKey,
			LegacyPhoneKey,
			CurrentPhoneKey,
			LegacyFaxKey,
			CurrentFaxKey
		};

		private static readonly Dictionary<string, string> LegacyToCurrent = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			[LegacyLawyerKey] = CurrentLawyerKey,
			[LegacyPostalCodeKey] = CurrentPostalCodeKey,
			[LegacyAddressKey] = CurrentAddressKey,
			[LegacyOfficeNameKey] = CurrentOfficeNameKey,
			[LegacyPhoneKey] = CurrentPhoneKey,
			[LegacyFaxKey] = CurrentFaxKey
		};

		private static readonly Dictionary<string, string> CurrentToLegacy = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			[CurrentLawyerKey] = LegacyLawyerKey,
			[CurrentPostalCodeKey] = LegacyPostalCodeKey,
			[CurrentAddressKey] = LegacyAddressKey,
			[CurrentOfficeNameKey] = LegacyOfficeNameKey,
			[CurrentPhoneKey] = LegacyPhoneKey,
			[CurrentFaxKey] = LegacyFaxKey
		};

		internal static string NormalizeToCurrent(string key)
		{
			string normalized = NormalizeKey(key);
			string current;
			return LegacyToCurrent.TryGetValue(normalized, out current) ? current : normalized;
		}

		internal static bool IsAllowedLegacyToCurrentRename(string oldKey, string newKey)
		{
			string normalizedOldKey = NormalizeKey(oldKey);
			string normalizedNewKey = NormalizeKey(newKey);
			string current;
			return LegacyToCurrent.TryGetValue(normalizedOldKey, out current)
				&& string.Equals(current, normalizedNewKey, StringComparison.OrdinalIgnoreCase);
		}

		internal static bool TryGetValueWithAliases(IReadOnlyDictionary<string, string> values, string key, out string value)
		{
			value = string.Empty;
			if (values == null)
			{
				return false;
			}

			string normalizedKey = NormalizeKey(key);
			if (normalizedKey.Length == 0)
			{
				return false;
			}

			if (values.TryGetValue(normalizedKey, out value))
			{
				value = value ?? string.Empty;
				return true;
			}

			string currentKey = NormalizeToCurrent(normalizedKey);
			if (!string.Equals(currentKey, normalizedKey, StringComparison.OrdinalIgnoreCase)
				&& values.TryGetValue(currentKey, out value))
			{
				value = value ?? string.Empty;
				return true;
			}

			string legacyKey;
			if (CurrentToLegacy.TryGetValue(normalizedKey, out legacyKey)
				&& values.TryGetValue(legacyKey, out value))
			{
				value = value ?? string.Empty;
				return true;
			}

			value = string.Empty;
			return false;
		}

		internal static string GetValueWithAliases(IReadOnlyDictionary<string, string> values, string key)
		{
			string value;
			return TryGetValueWithAliases(values, key, out value) ? value : string.Empty;
		}

		private static string NormalizeKey(string key)
		{
			return (key ?? string.Empty).Trim();
		}
	}
}
