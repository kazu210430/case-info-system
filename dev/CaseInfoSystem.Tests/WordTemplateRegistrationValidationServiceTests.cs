using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class WordTemplateRegistrationValidationServiceTests : IDisposable
	{
		private readonly List<string> _temporaryDirectories = new List<string> ();

		[Fact]
		public void Validate_WhenDocxAndDotxAreWellFormed_RegistersBothTemplates ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "01_委任状.docx"), BuildDocumentXml (CreateTextControl ("顧客_名前")));
			CreateWordPackage (Path.Combine (text, "12_通知書.dotx"), BuildDocumentXml (CreateRichTextControl ("Date")));

			TemplateRegistrationValidationSummary templateRegistrationValidationSummary = CreateService ().Validate (text, new string[1] { "顧客_名前" });

			Assert.Equal (2, templateRegistrationValidationSummary.DetectedFileCount);
			Assert.Equal (2, templateRegistrationValidationSummary.ValidTemplateCount);
			Assert.DoesNotContain (templateRegistrationValidationSummary.TemplateResults, entry => !entry.IsValid);
		}

		[Theory]
		[InlineData ("AA_委任状.docx")]
		[InlineData ("1_委任状.docx")]
		[InlineData ("001_委任状.docx")]
		[InlineData ("委任状.docx")]
		public void Validate_WhenFileNameDoesNotStartWithTwoDigitKey_ExcludesTemplate (string fileName)
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, fileName), BuildDocumentXml ());

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.False (templateRegistrationValidationEntry.IsValid);
			Assert.Contains ("ファイル名先頭が 2桁の key No. ではありません。", templateRegistrationValidationEntry.Errors);
		}

		[Theory]
		[InlineData ("00_委任状.docx")]
		[InlineData ("100_委任状.docx")]
		public void Validate_WhenKeyIsOutOfRange_ExcludesTemplate (string fileName)
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, fileName), BuildDocumentXml ());

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.False (templateRegistrationValidationEntry.IsValid);
			Assert.Contains ("key No. は 01～99 の範囲で指定してください。", templateRegistrationValidationEntry.Errors);
		}

		[Fact]
		public void Validate_WhenKeyIsDuplicated_ExcludesAllTemplatesForThatKey ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "08_報告書.docx"), BuildDocumentXml ());
			CreateWordPackage (Path.Combine (text, "08_通知書.docx"), BuildDocumentXml ());

			TemplateRegistrationValidationSummary templateRegistrationValidationSummary = CreateService ().Validate (text, Array.Empty<string> ());

			Assert.All (templateRegistrationValidationSummary.TemplateResults, entry => Assert.Contains ("key No. 08 が重複しています。", entry.Errors));
			Assert.Equal (0, templateRegistrationValidationSummary.ValidTemplateCount);
		}

		[Fact]
		public void Validate_WhenMacroEnabledTemplatesExist_ExcludesDocmAndDotm ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "21_契約書.docm"), BuildDocumentXml ());
			CreateWordPackage (Path.Combine (text, "22_案内.dotm"), BuildDocumentXml ());

			TemplateRegistrationValidationSummary templateRegistrationValidationSummary = CreateService ().Validate (text, Array.Empty<string> ());

			Assert.All (templateRegistrationValidationSummary.TemplateResults, entry => Assert.Contains ("マクロ付きファイルは登録できません。.docx または .dotx にしてください。", entry.Errors));
		}

		[Fact]
		public void Validate_WhenFileCannotBeReadAsWord_ExcludesTemplate ()
		{
			string text = CreateTempDirectory ();
			File.WriteAllText (Path.Combine (text, "31_破損.docx"), "not-a-word-package", Encoding.UTF8);

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.False (templateRegistrationValidationEntry.IsValid);
			Assert.Contains ("Wordファイルとして読み取れませんでした。", templateRegistrationValidationEntry.Errors);
		}

		[Fact]
		public void Validate_WhenUndefinedTextTagExists_ExcludesTemplate ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "41_通知書.docx"), BuildDocumentXml (CreateTextControl ("相手方_名前")));

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, new string[1] { "顧客_名前" }).TemplateResults);

			Assert.False (templateRegistrationValidationEntry.IsValid);
			Assert.Contains ("未定義タグ「相手方_名前」があります。", templateRegistrationValidationEntry.Errors);
		}

		[Fact]
		public void Validate_WhenDateTagExists_AllowsTemplate ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "42_受任通知.docx"), BuildDocumentXml (CreateTextControl ("Date")));

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.True (templateRegistrationValidationEntry.IsValid);
			Assert.Empty (templateRegistrationValidationEntry.Errors);
		}

		[Fact]
		public void Validate_WhenUndefinedNonTextTagExists_IgnoresControl ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "43_確認書.docx"), BuildDocumentXml (CreateCheckBoxControl ("未定義タグ")));

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.True (templateRegistrationValidationEntry.IsValid);
			Assert.Empty (templateRegistrationValidationEntry.Warnings);
		}

		[Fact]
		public void Validate_WhenTextTagIsBlank_AddsWarningButKeepsTemplate ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "44_照会書.docx"), BuildDocumentXml (CreateTextControl (null)));

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.True (templateRegistrationValidationEntry.IsValid);
			Assert.Contains ("Tag が未設定のテキスト項目があります。この項目は差し込み対象になりません。", templateRegistrationValidationEntry.Warnings);
		}

		[Fact]
		public void Validate_WhenTemplateHasNoTags_KeepsTemplate ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "45_固定文書.docx"), BuildDocumentXml ());

			TemplateRegistrationValidationEntry templateRegistrationValidationEntry = Assert.Single (CreateService ().Validate (text, Array.Empty<string> ()).TemplateResults);

			Assert.True (templateRegistrationValidationEntry.IsValid);
			Assert.Empty (templateRegistrationValidationEntry.Warnings);
		}

		[Fact]
		public void Validate_WhenMixedTemplatesPresent_OnlyValidTemplatesRemainRegistrableAndIssuesStayInResult ()
		{
			string text = CreateTempDirectory ();
			CreateWordPackage (Path.Combine (text, "01_委任状.docx"), BuildDocumentXml (CreateTextControl ("顧客_名前")));
			CreateWordPackage (Path.Combine (text, "05_照会書.docx"), BuildDocumentXml (CreateTextControl (null)));
			CreateWordPackage (Path.Combine (text, "08_報告書.docx"), BuildDocumentXml ());
			CreateWordPackage (Path.Combine (text, "08_通知書.docx"), BuildDocumentXml ());
			CreateWordPackage (Path.Combine (text, "12_通知書.docm"), BuildDocumentXml ());
			CreateWordPackage (Path.Combine (text, "21_契約書.docx"), BuildDocumentXml (CreateTextControl ("相手方_名前")));

			TemplateRegistrationValidationSummary templateRegistrationValidationSummary = CreateService ().Validate (text, new string[1] { "顧客_名前" });
			IReadOnlyList<TemplateRegistrationValidationEntry> readOnlyList = templateRegistrationValidationSummary.GetValidTemplates ();

			Assert.Equal (6, templateRegistrationValidationSummary.DetectedFileCount);
			Assert.Equal (2, readOnlyList.Count);
			Assert.Equal (new string[2] { "01", "05" }, readOnlyList.Select (entry => entry.Key).ToArray ());
			Assert.Contains (templateRegistrationValidationSummary.TemplateResults, entry => entry.FileName == "21_契約書.docx" && entry.Errors.Contains ("未定義タグ「相手方_名前」があります。"));
			Assert.Contains (templateRegistrationValidationSummary.TemplateResults, entry => entry.FileName == "05_照会書.docx" && entry.Warnings.Contains ("Tag が未設定のテキスト項目があります。この項目は差し込み対象になりません。"));
		}

		public void Dispose ()
		{
			foreach (string temporaryDirectory in _temporaryDirectories) {
				try {
					if (Directory.Exists (temporaryDirectory)) {
						Directory.Delete (temporaryDirectory, recursive: true);
					}
				} catch {
				}
			}
		}

		private static WordTemplateRegistrationValidationService CreateService ()
		{
			Logger logger = OrchestrationTestSupport.CreateLogger (new List<string> ());
			return new WordTemplateRegistrationValidationService (new WordTemplateContentControlInspectionService (), logger);
		}

		private string CreateTempDirectory ()
		{
			string path = Path.Combine (Path.GetTempPath (), "CaseInfoSystem.TemplateValidation." + Guid.NewGuid ().ToString ("N"));
			Directory.CreateDirectory (path);
			_temporaryDirectories.Add (path);
			return path;
		}

		private static void CreateWordPackage (string fullPath, string documentXml)
		{
			using (ZipArchive zipArchive = ZipFile.Open (fullPath, ZipArchiveMode.Create)) {
				ZipArchiveEntry zipArchiveEntry = zipArchive.CreateEntry ("word/document.xml");
				using (StreamWriter streamWriter = new StreamWriter (zipArchiveEntry.Open (), new UTF8Encoding (encoderShouldEmitUTF8Identifier: false))) {
					streamWriter.Write (documentXml);
				}
			}
		}

		private static string BuildDocumentXml (params string[] contentControls)
		{
			return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
				+ "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
				+ "<w:body>"
				+ string.Concat (contentControls ?? Array.Empty<string> ())
				+ "<w:sectPr/>"
				+ "</w:body>"
				+ "</w:document>";
		}

		private static string CreateTextControl (string tag)
		{
			return "<w:sdt><w:sdtPr>" + BuildTagElement (tag) + "<w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>sample</w:t></w:r></w:p></w:sdtContent></w:sdt>";
		}

		private static string CreateRichTextControl (string tag)
		{
			return "<w:sdt><w:sdtPr>" + BuildTagElement (tag) + "<w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>sample</w:t></w:r></w:p></w:sdtContent></w:sdt>";
		}

		private static string CreateCheckBoxControl (string tag)
		{
			return "<w:sdt><w:sdtPr>" + BuildTagElement (tag) + "<w14:checkbox xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>sample</w:t></w:r></w:p></w:sdtContent></w:sdt>";
		}

		private static string BuildTagElement (string tag)
		{
			return (tag == null) ? string.Empty : "<w:tag w:val=\"" + SecurityElement.Escape (tag) + "\"/>";
		}
	}
}
