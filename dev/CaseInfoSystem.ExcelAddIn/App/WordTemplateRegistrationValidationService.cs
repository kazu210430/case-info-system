using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class WordTemplateRegistrationValidationService
	{
		private static readonly string[] CandidateExtensions = new string[4] { ".docx", ".dotx", ".docm", ".dotm" };

		private const string EmptyTextTagWarningMessage = "Tag が未設定のテキスト項目があります。この項目は差し込み対象になりません。";

		private const string MacroEnabledFileErrorMessage = "マクロ付きファイルは登録できません。.docx または .dotx にしてください。";

		private const string InvalidKeyFormatErrorMessage = "ファイル名先頭が 2桁の key No. ではありません。";

		private const string InvalidKeyRangeErrorMessage = "key No. は 01～99 の範囲で指定してください。";

		private const string UnreadableWordFileErrorMessage = "Wordファイルとして読み取れませんでした。";

		private const string AllowedDateTag = "Date";

		private readonly WordTemplateContentControlInspectionService _inspectionService;

		private readonly Logger _logger;

		internal WordTemplateRegistrationValidationService (WordTemplateContentControlInspectionService inspectionService, Logger logger)
		{
			_inspectionService = inspectionService ?? throw new ArgumentNullException ("inspectionService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal TemplateRegistrationValidationSummary Validate (string templateDirectory, IEnumerable<string> definedTags)
		{
			if (string.IsNullOrWhiteSpace (templateDirectory)) {
				throw new ArgumentNullException ("templateDirectory");
			}
			HashSet<string> hashSet = BuildDefinedTagSet (definedTags);
			TemplateRegistrationValidationSummary templateRegistrationValidationSummary = new TemplateRegistrationValidationSummary {
				TemplateDirectory = templateDirectory
			};
			IReadOnlyList<string> candidateFiles = EnumerateCandidateFiles (templateDirectory);
			templateRegistrationValidationSummary.DetectedFileCount = candidateFiles.Count;
			foreach (string item in candidateFiles) {
				templateRegistrationValidationSummary.TemplateResults.Add (ValidateSingleTemplate (item, hashSet));
			}
			ApplyDuplicateKeyErrors (templateRegistrationValidationSummary.TemplateResults);
			return templateRegistrationValidationSummary;
		}

		private TemplateRegistrationValidationEntry ValidateSingleTemplate (string templatePath, ISet<string> definedTags)
		{
			string fileName = Path.GetFileName (templatePath) ?? string.Empty;
			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = new TemplateRegistrationValidationEntry {
				FileName = fileName,
				DisplayName = ExtractDisplayName (fileName)
			};
			ValidateFileExtension (templateRegistrationValidationEntry, fileName);
			ValidateKey (templateRegistrationValidationEntry, fileName);
			if (IsSupportedOpenXmlTemplate (fileName)) {
				InspectTextContentControls (templateRegistrationValidationEntry, templatePath, definedTags);
			}
			return templateRegistrationValidationEntry;
		}

		private static HashSet<string> BuildDefinedTagSet (IEnumerable<string> definedTags)
		{
			HashSet<string> hashSet = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
			if (definedTags == null) {
				return hashSet;
			}
			foreach (string definedTag in definedTags) {
				string text = (definedTag ?? string.Empty).Trim ();
				if (text.Length != 0) {
					hashSet.Add (text);
				}
			}
			return hashSet;
		}

		private static IReadOnlyList<string> EnumerateCandidateFiles (string templateDirectory)
		{
			string[] files = Directory.GetFiles (templateDirectory, "*.*", SearchOption.TopDirectoryOnly);
			Array.Sort (files, StringComparer.OrdinalIgnoreCase);
			return files
				.Where (path => IsCandidateExtension (Path.GetExtension (path) ?? string.Empty))
				.Where (path => !IsOfficeTemporaryFile (Path.GetFileName (path) ?? string.Empty))
				.ToList ();
		}

		private static bool IsCandidateExtension (string extension)
		{
			return CandidateExtensions.Any (candidateExtension => string.Equals (candidateExtension, extension, StringComparison.OrdinalIgnoreCase));
		}

		private static bool IsOfficeTemporaryFile (string fileName)
		{
			return !string.IsNullOrWhiteSpace (fileName) && fileName.StartsWith ("~$", StringComparison.OrdinalIgnoreCase);
		}

		private static void ValidateFileExtension (TemplateRegistrationValidationEntry entry, string fileName)
		{
			if (entry == null) {
				throw new ArgumentNullException ("entry");
			}
			string extension = Path.GetExtension (fileName) ?? string.Empty;
			if (string.Equals (extension, ".docm", StringComparison.OrdinalIgnoreCase)
				|| string.Equals (extension, ".dotm", StringComparison.OrdinalIgnoreCase)) {
				entry.AddError (MacroEnabledFileErrorMessage);
			}
		}

		private static void ValidateKey (TemplateRegistrationValidationEntry entry, string fileName)
		{
			if (entry == null) {
				throw new ArgumentNullException ("entry");
			}
			string text = Path.GetFileNameWithoutExtension (fileName) ?? string.Empty;
			int num = text.IndexOf ('_');
			if (num <= 0) {
				entry.AddError (InvalidKeyFormatErrorMessage);
				return;
			}
			string text2 = text.Substring (0, num);
			if (text2.Length != 2 || !text2.All (char.IsDigit)) {
				entry.AddError (InvalidKeyFormatErrorMessage);
				if (text2.All (char.IsDigit) && int.TryParse (text2, out var result) && (result <= 0 || result >= 100)) {
					entry.AddError (InvalidKeyRangeErrorMessage);
				}
				return;
			}
			if (!int.TryParse (text2, out var result2) || result2 < 1 || result2 > 99) {
				entry.AddError (InvalidKeyRangeErrorMessage);
				return;
			}
			entry.Key = result2.ToString ("00");
		}

		private void InspectTextContentControls (TemplateRegistrationValidationEntry entry, string templatePath, ISet<string> definedTags)
		{
			if (entry == null) {
				throw new ArgumentNullException ("entry");
			}
			try {
				WordTemplateContentControlInspectionService.InspectionResult inspectionResult = _inspectionService.Inspect (templatePath);
				if (inspectionResult.EmptyTextTagCount > 0) {
					entry.AddWarning (EmptyTextTagWarningMessage);
				}
				foreach (string item in inspectionResult.TextTags.OrderBy (tag => tag, StringComparer.OrdinalIgnoreCase)) {
					if (!definedTags.Contains (item) && !string.Equals (item, AllowedDateTag, StringComparison.OrdinalIgnoreCase)) {
						entry.AddError ("未定義タグ「" + item + "」があります。");
					}
				}
			} catch (Exception ex) {
				_logger.Warn ("WordTemplateRegistrationValidationService inspection failed. file=" + (entry.FileName ?? string.Empty) + ", error=" + ex.Message);
				entry.AddError (UnreadableWordFileErrorMessage);
			}
		}

		private static void ApplyDuplicateKeyErrors (IEnumerable<TemplateRegistrationValidationEntry> entries)
		{
			Dictionary<string, List<TemplateRegistrationValidationEntry>> dictionary = new Dictionary<string, List<TemplateRegistrationValidationEntry>> (StringComparer.OrdinalIgnoreCase);
			if (entries == null) {
				return;
			}
			foreach (TemplateRegistrationValidationEntry entry in entries) {
				string text = (entry == null) ? string.Empty : (entry.Key ?? string.Empty).Trim ();
				if (text.Length == 0) {
					continue;
				}
				if (!dictionary.TryGetValue (text, out var value)) {
					value = new List<TemplateRegistrationValidationEntry> ();
					dictionary.Add (text, value);
				}
				value.Add (entry);
			}
			foreach (KeyValuePair<string, List<TemplateRegistrationValidationEntry>> item in dictionary) {
				if (item.Value.Count <= 1) {
					continue;
				}
				string message = "key No. " + item.Key + " が重複しています。";
				foreach (TemplateRegistrationValidationEntry item2 in item.Value) {
					item2.AddError (message);
				}
			}
		}

		private static bool IsSupportedOpenXmlTemplate (string fileName)
		{
			string extension = Path.GetExtension (fileName) ?? string.Empty;
			return string.Equals (extension, ".docx", StringComparison.OrdinalIgnoreCase)
				|| string.Equals (extension, ".dotx", StringComparison.OrdinalIgnoreCase);
		}

		private static string ExtractDisplayName (string fileName)
		{
			string text = Path.GetFileNameWithoutExtension (fileName) ?? string.Empty;
			int num = text.IndexOf ('_');
			if (num >= 0 && num + 1 < text.Length) {
				return text.Substring (num + 1);
			}
			return text;
		}
	}
}
