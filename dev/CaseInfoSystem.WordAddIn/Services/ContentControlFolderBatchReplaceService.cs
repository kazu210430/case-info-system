using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace CaseInfoSystem.WordAddIn.Services
{
	internal sealed class ContentControlFolderBatchReplaceService
	{
		internal sealed class FolderReplaceRequest
		{
			public string TemplateDirectory { get; set; }

			public string OldTag { get; set; }

			public string NewTag { get; set; }

			public string OldTitle { get; set; }

			public string NewTitle { get; set; }

			public string OldDisplayText { get; set; }

			public string NewDisplayText { get; set; }

			public bool UsePartialMatch { get; set; }

			public bool CreateBackups { get; set; }
		}

		internal sealed class FolderReplaceResult
		{
			private readonly List<FileReplaceResult> _fileResults = new List<FileReplaceResult> ();

			public string TemplateDirectory { get; set; }

			public IReadOnlyList<FileReplaceResult> FileResults
			{
				get { return _fileResults; }
			}

			public int DetectedFileCount { get; private set; }

			public int ProcessedFileCount { get; private set; }

			public int ChangedFileCount { get; private set; }

			public int FailedFileCount { get; private set; }

			public int BackupCreatedCount { get; private set; }

			public int ScannedCount { get; private set; }

			public int TagChangedCount { get; private set; }

			public int TitleChangedCount { get; private set; }

			public int DisplayTextChangedCount { get; private set; }

			internal void AddFileResult (FileReplaceResult result)
			{
				if (result == null) {
					return;
				}

				_fileResults.Add (result);
				DetectedFileCount++;
				if (result.Success) {
					ProcessedFileCount++;
					ScannedCount += result.ScannedCount;
					TagChangedCount += result.TagChangedCount;
					TitleChangedCount += result.TitleChangedCount;
					DisplayTextChangedCount += result.DisplayTextChangedCount;
					if (result.HasChanges) {
						ChangedFileCount++;
					}
					if (result.BackupCreated) {
						BackupCreatedCount++;
					}
					return;
				}

				FailedFileCount++;
			}
		}

		internal sealed class FileReplaceResult
		{
			public string FilePath { get; set; }

			public string FileName { get; set; }

			public bool Success { get; set; }

			public string FailureMessage { get; set; }

			public int ScannedCount { get; set; }

			public int TagChangedCount { get; set; }

			public int TitleChangedCount { get; set; }

			public int DisplayTextChangedCount { get; set; }

			public bool BackupCreated { get; set; }

			public string BackupPath { get; set; }

			public bool HasChanges
			{
				get { return TagChangedCount > 0 || TitleChangedCount > 0 || DisplayTextChangedCount > 0; }
			}
		}

		private sealed class ChangedEntry
		{
			internal ChangedEntry (string fullName, string content)
			{
				FullName = fullName ?? string.Empty;
				Content = content ?? string.Empty;
			}

			internal string FullName { get; private set; }

			internal string Content { get; private set; }
		}

		private static readonly XNamespace WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		private static readonly string[] SupportedExtensions = new string[] {
			".docx",
			".dotx",
			".docm",
			".dotm"
		};

		internal FolderReplaceResult Execute (FolderReplaceRequest request)
		{
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			if (string.IsNullOrWhiteSpace (request.TemplateDirectory)) {
				throw new ArgumentException ("雛形フォルダを指定してください。", "request");
			}
			if (!Directory.Exists (request.TemplateDirectory)) {
				throw new DirectoryNotFoundException ("雛形フォルダが見つかりません: " + request.TemplateDirectory);
			}

			FolderReplaceResult folderReplaceResult = new FolderReplaceResult {
				TemplateDirectory = request.TemplateDirectory
			};
			foreach (string filePath in EnumerateCandidateFiles (request.TemplateDirectory)) {
				folderReplaceResult.AddFileResult (ProcessFile (filePath, request));
			}
			return folderReplaceResult;
		}

		internal static bool HasAnyTarget (FolderReplaceRequest request)
		{
			if (request == null) {
				return false;
			}
			return !string.IsNullOrEmpty (request.OldTag)
				|| !string.IsNullOrEmpty (request.OldTitle)
				|| !string.IsNullOrEmpty (request.OldDisplayText);
		}

		internal static string BuildCompletionMessage (FolderReplaceResult result)
		{
			if (result == null) {
				return "フォルダ一括置換の結果を取得できませんでした。";
			}

			StringBuilder stringBuilder = new StringBuilder ();
			stringBuilder.AppendLine ("雛形フォルダ内のコンテンツコントロール置換が完了しました。");
			stringBuilder.AppendLine ("対象フォルダ: " + (result.TemplateDirectory ?? string.Empty));
			stringBuilder.AppendLine ("検出ファイル数: " + result.DetectedFileCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("処理成功ファイル数: " + result.ProcessedFileCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("変更ファイル数: " + result.ChangedFileCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("バックアップ作成数: " + result.BackupCreatedCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("処理失敗ファイル数: " + result.FailedFileCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("対象コントロール数: " + result.ScannedCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("Tag 変更数: " + result.TagChangedCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.AppendLine ("Title 変更数: " + result.TitleChangedCount.ToString (CultureInfo.InvariantCulture));
			stringBuilder.Append ("表示文字 変更数: " + result.DisplayTextChangedCount.ToString (CultureInfo.InvariantCulture));
			if (result.FailedFileCount > 0) {
				stringBuilder.AppendLine ();
				stringBuilder.AppendLine ();
				stringBuilder.AppendLine ("失敗したファイル:");
				foreach (FileReplaceResult fileResult in result.FileResults.Where (item => !item.Success)) {
					stringBuilder.AppendLine ("- " + fileResult.FileName + ": " + fileResult.FailureMessage);
				}
			}
			return stringBuilder.ToString ();
		}

		private static IReadOnlyList<string> EnumerateCandidateFiles (string templateDirectory)
		{
			return Directory.GetFiles (templateDirectory, "*.*", SearchOption.TopDirectoryOnly)
				.Where (IsSupportedTemplateFile)
				.OrderBy (Path.GetFileName, StringComparer.OrdinalIgnoreCase)
				.ToList ();
		}

		private static bool IsSupportedTemplateFile (string filePath)
		{
			string extension = Path.GetExtension (filePath) ?? string.Empty;
			return SupportedExtensions.Any (item => string.Equals (item, extension, StringComparison.OrdinalIgnoreCase));
		}

		private static FileReplaceResult ProcessFile (string filePath, FolderReplaceRequest request)
		{
			FileReplaceResult fileReplaceResult = new FileReplaceResult {
				FilePath = filePath,
				FileName = Path.GetFileName (filePath) ?? string.Empty
			};

			try {
				List<ChangedEntry> changedEntries = BuildChangedEntries (filePath, request, fileReplaceResult);
				if (changedEntries.Count > 0) {
					if (request.CreateBackups) {
						fileReplaceResult.BackupPath = CreateBackup (filePath);
						fileReplaceResult.BackupCreated = true;
					}
					WriteChangedEntries (filePath, changedEntries);
				}

				fileReplaceResult.Success = true;
			} catch (Exception exception) {
				fileReplaceResult.Success = false;
				fileReplaceResult.FailureMessage = exception.Message;
			}

			return fileReplaceResult;
		}

		private static List<ChangedEntry> BuildChangedEntries (string filePath, FolderReplaceRequest request, FileReplaceResult result)
		{
			List<ChangedEntry> changedEntries = new List<ChangedEntry> ();
			using (ZipArchive archive = ZipFile.OpenRead (filePath)) {
				EnsureMainDocumentExists (archive);
				foreach (ZipArchiveEntry entry in archive.Entries) {
					if (!ShouldInspectEntry (entry)) {
						continue;
					}

					XDocument document;
					using (Stream stream = entry.Open ()) {
						document = XDocument.Load (stream, LoadOptions.PreserveWhitespace);
					}

					bool changed = ApplyReplacement (document, request, result);
					if (changed) {
						changedEntries.Add (new ChangedEntry (entry.FullName, Serialize (document)));
					}
				}
			}
			return changedEntries;
		}

		private static void WriteChangedEntries (string filePath, IEnumerable<ChangedEntry> changedEntries)
		{
			using (ZipArchive archive = ZipFile.Open (filePath, ZipArchiveMode.Update)) {
				foreach (ChangedEntry changedEntry in changedEntries) {
					ZipArchiveEntry existingEntry = archive.GetEntry (changedEntry.FullName);
					if (existingEntry != null) {
						existingEntry.Delete ();
					}

					ZipArchiveEntry newEntry = archive.CreateEntry (changedEntry.FullName, CompressionLevel.Optimal);
					using (Stream stream = newEntry.Open ())
					using (StreamWriter writer = new StreamWriter (stream, new UTF8Encoding (false))) {
						writer.Write (changedEntry.Content);
					}
				}
			}
		}

		private static string CreateBackup (string filePath)
		{
			string backupPath = filePath + ".bak";
			if (!File.Exists (backupPath)) {
				File.Copy (filePath, backupPath, false);
				return backupPath;
			}

			string timestamp = DateTime.Now.ToString ("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
			for (int index = 1; index < 1000; index++) {
				string candidatePath = filePath + "." + timestamp + "." + index.ToString (CultureInfo.InvariantCulture) + ".bak";
				if (File.Exists (candidatePath)) {
					continue;
				}
				File.Copy (filePath, candidatePath, false);
				return candidatePath;
			}

			throw new IOException ("バックアップファイル名を決定できませんでした。");
		}

		private static bool ApplyReplacement (XDocument document, FolderReplaceRequest request, FileReplaceResult result)
		{
			bool changed = false;
			foreach (XElement contentControl in document.Descendants (WordNamespace + "sdt")) {
				XElement propertiesElement = contentControl.Element (WordNamespace + "sdtPr");
				if (propertiesElement == null || !IsTextContentControl (propertiesElement)) {
					continue;
				}

				result.ScannedCount++;
				if (ReplaceValAttribute (propertiesElement.Element (WordNamespace + "tag"), request.OldTag, request.NewTag, request.UsePartialMatch)) {
					result.TagChangedCount++;
					changed = true;
				}
				if (ReplaceValAttribute (propertiesElement.Element (WordNamespace + "alias"), request.OldTitle, request.NewTitle, request.UsePartialMatch)) {
					result.TitleChangedCount++;
					changed = true;
				}
				if (ReplaceDisplayText (contentControl, request.OldDisplayText, request.NewDisplayText, request.UsePartialMatch)) {
					result.DisplayTextChangedCount++;
					changed = true;
				}
			}
			return changed;
		}

		private static bool IsTextContentControl (XElement propertiesElement)
		{
			return propertiesElement.Element (WordNamespace + "text") != null
				|| propertiesElement.Element (WordNamespace + "richText") != null;
		}

		private static bool ReplaceValAttribute (XElement element, string oldValue, string newValue, bool usePartialMatch)
		{
			if (element == null || string.IsNullOrEmpty (oldValue)) {
				return false;
			}

			XAttribute valueAttribute = element.Attribute (WordNamespace + "val");
			string currentValue = valueAttribute == null ? string.Empty : (valueAttribute.Value ?? string.Empty);
			string replacedValue = ReplaceValue (currentValue, oldValue, newValue, usePartialMatch);
			if (string.Equals (replacedValue, currentValue, StringComparison.Ordinal)) {
				return false;
			}

			element.SetAttributeValue (WordNamespace + "val", replacedValue);
			return true;
		}

		private static string ReplaceValue (string source, string oldValue, string newValue, bool usePartialMatch)
		{
			if (string.IsNullOrEmpty (oldValue)) {
				return source ?? string.Empty;
			}
			string text = source ?? string.Empty;
			string text2 = newValue ?? string.Empty;
			if (usePartialMatch) {
				return text.Replace (oldValue, text2);
			}
			return string.Equals (text, oldValue, StringComparison.Ordinal) ? text2 : text;
		}

		private static bool ReplaceDisplayText (XElement contentControl, string oldValue, string newValue, bool usePartialMatch)
		{
			if (contentControl == null || string.IsNullOrEmpty (oldValue)) {
				return false;
			}

			XElement contentElement = contentControl.Element (WordNamespace + "sdtContent");
			if (contentElement == null) {
				return false;
			}

			List<XElement> textElements = contentElement.Descendants (WordNamespace + "t").ToList ();
			if (textElements.Count == 0) {
				return false;
			}

			string currentDisplayText = string.Concat (textElements.Select (item => item.Value ?? string.Empty));
			string replacedDisplayText = ReplaceValue (currentDisplayText, oldValue, newValue, usePartialMatch);
			if (string.Equals (replacedDisplayText, currentDisplayText, StringComparison.Ordinal)
				&& !TryReplaceDisplayTextIgnoringLayoutWhitespace (currentDisplayText, oldValue, newValue, usePartialMatch, out replacedDisplayText)) {
				return false;
			}

			WriteDisplayText (textElements, replacedDisplayText);
			return true;
		}

		private static bool TryReplaceDisplayTextIgnoringLayoutWhitespace (
			string source,
			string oldValue,
			string newValue,
			bool usePartialMatch,
			out string replacedText)
		{
			replacedText = source ?? string.Empty;
			if (string.IsNullOrEmpty (source) || string.IsNullOrEmpty (oldValue)) {
				return false;
			}

			CompactTextMap sourceMap = CreateCompactTextMap (source);
			CompactTextMap oldValueMap = CreateCompactTextMap (oldValue);
			if (string.IsNullOrEmpty (oldValueMap.Text) || string.IsNullOrEmpty (sourceMap.Text)) {
				return false;
			}

			if (!usePartialMatch) {
				if (!string.Equals (sourceMap.Text, oldValueMap.Text, StringComparison.Ordinal)) {
					return false;
				}

				replacedText = FormatReplacementForRawSpan (source, newValue);
				return !string.Equals (replacedText, source, StringComparison.Ordinal);
			}

			List<RawReplacement> replacements = new List<RawReplacement> ();
			int searchIndex = 0;
			while (searchIndex <= sourceMap.Text.Length - oldValueMap.Text.Length) {
				int matchIndex = sourceMap.Text.IndexOf (oldValueMap.Text, searchIndex, StringComparison.Ordinal);
				if (matchIndex < 0) {
					break;
				}

				int rawStart = sourceMap.SourceIndexes[matchIndex];
				int rawEnd = sourceMap.SourceIndexes[matchIndex + oldValueMap.Text.Length - 1] + 1;
				string rawSpan = source.Substring (rawStart, rawEnd - rawStart);
				string replacement = FormatReplacementForRawSpan (rawSpan, newValue);
				replacements.Add (new RawReplacement (rawStart, rawEnd, replacement));
				searchIndex = matchIndex + oldValueMap.Text.Length;
			}

			if (replacements.Count == 0) {
				return false;
			}

			StringBuilder stringBuilder = new StringBuilder (source);
			for (int index = replacements.Count - 1; index >= 0; index--) {
				RawReplacement replacement = replacements[index];
				stringBuilder.Remove (replacement.StartIndex, replacement.EndIndex - replacement.StartIndex);
				stringBuilder.Insert (replacement.StartIndex, replacement.Value);
			}

			replacedText = stringBuilder.ToString ();
			return !string.Equals (replacedText, source, StringComparison.Ordinal);
		}

		private static CompactTextMap CreateCompactTextMap (string source)
		{
			StringBuilder stringBuilder = new StringBuilder ();
			List<int> sourceIndexes = new List<int> ();
			string text = source ?? string.Empty;
			for (int index = 0; index < text.Length; index++) {
				char character = text[index];
				if (IsLayoutWhitespace (character)) {
					continue;
				}

				stringBuilder.Append (character);
				sourceIndexes.Add (index);
			}

			return new CompactTextMap (stringBuilder.ToString (), sourceIndexes);
		}

		private static string FormatReplacementForRawSpan (string rawSpan, string newValue)
		{
			string replacement = newValue ?? string.Empty;
			string separator = FindInterCharacterLayoutSeparator (rawSpan);
			if (string.IsNullOrEmpty (separator)) {
				return replacement;
			}

			string compactReplacement = CreateCompactTextMap (replacement).Text;
			if (compactReplacement.Length <= 1) {
				return compactReplacement;
			}

			StringBuilder stringBuilder = new StringBuilder ();
			for (int index = 0; index < compactReplacement.Length; index++) {
				if (index > 0) {
					stringBuilder.Append (separator);
				}
				stringBuilder.Append (compactReplacement[index]);
			}
			return stringBuilder.ToString ();
		}

		private static string FindInterCharacterLayoutSeparator (string rawSpan)
		{
			string text = rawSpan ?? string.Empty;
			bool hasPreviousText = false;
			for (int index = 0; index < text.Length;) {
				if (!IsLayoutWhitespace (text[index])) {
					hasPreviousText = true;
					index++;
					continue;
				}

				int startIndex = index;
				while (index < text.Length && IsLayoutWhitespace (text[index])) {
					index++;
				}

				if (hasPreviousText && index < text.Length) {
					return text.Substring (startIndex, index - startIndex);
				}
			}

			return string.Empty;
		}

		private static bool IsLayoutWhitespace (char character)
		{
			return char.IsWhiteSpace (character) || character == '\u00A0';
		}

		private static void WriteDisplayText (IReadOnlyList<XElement> textElements, string newValue)
		{
			for (int index = 0; index < textElements.Count; index++) {
				SetTextElementValue (textElements[index], index == 0 ? newValue : string.Empty);
			}
		}

		private static void SetTextElementValue (XElement textElement, string value)
		{
			string text = value ?? string.Empty;
			textElement.Value = text;
			XNamespace xmlNamespace = XNamespace.Xml;
			if (NeedsPreserveSpace (text)) {
				textElement.SetAttributeValue (xmlNamespace + "space", "preserve");
				return;
			}
			textElement.SetAttributeValue (xmlNamespace + "space", null);
		}

		private static bool NeedsPreserveSpace (string text)
		{
			if (string.IsNullOrEmpty (text)) {
				return false;
			}
			return char.IsWhiteSpace (text[0]) || char.IsWhiteSpace (text[text.Length - 1]);
		}

		private static void EnsureMainDocumentExists (ZipArchive archive)
		{
			if (archive.GetEntry ("word/document.xml") == null) {
				throw new InvalidDataException ("word/document.xml が見つかりません。");
			}
		}

		private static bool ShouldInspectEntry (ZipArchiveEntry entry)
		{
			if (entry == null) {
				return false;
			}
			string fullName = entry.FullName ?? string.Empty;
			return fullName.StartsWith ("word/", StringComparison.OrdinalIgnoreCase)
				&& fullName.EndsWith (".xml", StringComparison.OrdinalIgnoreCase);
		}

		private static string Serialize (XDocument document)
		{
			XmlWriterSettings settings = new XmlWriterSettings {
				Encoding = new UTF8Encoding (false),
				Indent = false,
				OmitXmlDeclaration = document.Declaration == null
			};
			using (StringWriter stringWriter = new Utf8StringWriter ())
			using (XmlWriter xmlWriter = XmlWriter.Create (stringWriter, settings)) {
				document.Save (xmlWriter);
				xmlWriter.Flush ();
				return stringWriter.ToString ();
			}
		}

		private sealed class Utf8StringWriter : StringWriter
		{
			public override Encoding Encoding
			{
				get { return Encoding.UTF8; }
			}
		}

		private sealed class CompactTextMap
		{
			public CompactTextMap (string text, List<int> sourceIndexes)
			{
				Text = text ?? string.Empty;
				SourceIndexes = sourceIndexes ?? new List<int> ();
			}

			public string Text { get; private set; }

			public List<int> SourceIndexes { get; private set; }
		}

		private sealed class RawReplacement
		{
			public RawReplacement (int startIndex, int endIndex, string value)
			{
				StartIndex = startIndex;
				EndIndex = endIndex;
				Value = value ?? string.Empty;
			}

			public int StartIndex { get; private set; }

			public int EndIndex { get; private set; }

			public string Value { get; private set; }
		}
	}
}
