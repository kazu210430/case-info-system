using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class WordTemplateContentControlInspectionService
	{
		internal sealed class InspectionResult
		{
			private readonly HashSet<string> _textTags = new HashSet<string> (StringComparer.OrdinalIgnoreCase);

			internal IReadOnlyCollection<string> TextTags => _textTags;

			internal int EmptyTextTagCount { get; private set; }

			internal void AddTextTag (string tag)
			{
				string text = (tag ?? string.Empty).Trim ();
				if (text.Length == 0) {
					EmptyTextTagCount++;
					return;
				}
				_textTags.Add (text);
			}
		}

		private static readonly XNamespace WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		internal InspectionResult Inspect (string templatePath)
		{
			if (string.IsNullOrWhiteSpace (templatePath)) {
				throw new ArgumentNullException ("templatePath");
			}
			InspectionResult inspectionResult = new InspectionResult ();
			using (ZipArchive zipArchive = ZipFile.OpenRead (templatePath)) {
				EnsureMainDocumentExists (zipArchive);
				foreach (ZipArchiveEntry item in zipArchive.Entries) {
					if (!ShouldInspectEntry (item)) {
						continue;
					}
					using (Stream stream = item.Open ()) {
						XDocument xDocument = XDocument.Load (stream, LoadOptions.None);
						foreach (XElement item2 in xDocument.Descendants (WordNamespace + "sdt")) {
							XElement xElement = item2.Element (WordNamespace + "sdtPr");
							if (xElement != null && IsTextContentControl (xElement)) {
								inspectionResult.AddTextTag (ReadTagValue (xElement));
							}
						}
					}
				}
			}
			return inspectionResult;
		}

		private static void EnsureMainDocumentExists (ZipArchive archive)
		{
			if (archive == null) {
				throw new ArgumentNullException ("archive");
			}
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

		private static bool IsTextContentControl (XElement propertiesElement)
		{
			if (propertiesElement == null) {
				return false;
			}
			return propertiesElement.Element (WordNamespace + "text") != null
				|| propertiesElement.Element (WordNamespace + "richText") != null;
		}

		private static string ReadTagValue (XElement propertiesElement)
		{
			if (propertiesElement == null) {
				return string.Empty;
			}
			XElement xElement = propertiesElement.Element (WordNamespace + "tag");
			XAttribute xAttribute = xElement?.Attribute (WordNamespace + "val");
			return (xAttribute == null) ? string.Empty : (xAttribute.Value ?? string.Empty);
		}
	}
}
