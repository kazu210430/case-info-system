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
	public sealed class KernelTemplateSyncPreflightServiceTests : IDisposable
	{
		private readonly List<string> _temporaryDirectories = new List<string> ();

		[Fact]
		public void Run_WhenDefinedTemplateTagsAreMissing_ReturnsFailedPreflight ()
		{
			string text = CreateTempSystemRoot ();
			CreateWordPackage (Path.Combine (GetTemplateDirectory (text), "01_委任状.docx"), BuildDocumentXml (CreateTextControl ("顧客_名前")));

			KernelTemplateSyncPreflightResult kernelTemplateSyncPreflightResult = CreateService ().Run (new KernelTemplateSyncPreflightRequest (text, Array.Empty<string> ()));

			Assert.Equal (KernelTemplateSyncPreflightStatus.Failed, kernelTemplateSyncPreflightResult.Status);
			Assert.Equal (GetTemplateDirectory (text), kernelTemplateSyncPreflightResult.TemplateDirectory);
			Assert.Null (kernelTemplateSyncPreflightResult.ValidationSummary);
			Assert.NotNull (kernelTemplateSyncPreflightResult.Failure);
			Assert.Equal (ValidationFailureKind.MissingDefinedTemplateTags, kernelTemplateSyncPreflightResult.Failure.Kind);
			Assert.Equal ("Kernelブックの管理シート CaseList_FieldInventory を読み取れません。", kernelTemplateSyncPreflightResult.Failure.Message);
			Assert.Empty (kernelTemplateSyncPreflightResult.Failure.TemplateResults);
		}

		[Fact]
		public void Run_WhenNoCandidateFilesAreFound_ReturnsFailedPreflightWithValidationSummary ()
		{
			string text = CreateTempSystemRoot ();

			KernelTemplateSyncPreflightResult kernelTemplateSyncPreflightResult = CreateService ().Run (new KernelTemplateSyncPreflightRequest (text, new string[1] { "顧客_名前" }));

			Assert.Equal (KernelTemplateSyncPreflightStatus.Failed, kernelTemplateSyncPreflightResult.Status);
			Assert.NotNull (kernelTemplateSyncPreflightResult.ValidationSummary);
			Assert.Equal (0, kernelTemplateSyncPreflightResult.ValidationSummary.DetectedFileCount);
			Assert.NotNull (kernelTemplateSyncPreflightResult.Failure);
			Assert.Equal (ValidationFailureKind.NoTemplateFiles, kernelTemplateSyncPreflightResult.Failure.Kind);
			Assert.Equal (0, kernelTemplateSyncPreflightResult.Failure.DetectedCount);
			Assert.Contains ("雛形フォルダに Word 雛形", kernelTemplateSyncPreflightResult.Failure.Message);
		}

		[Fact]
		public void Run_WhenTemplatesAreDetected_ReturnsSucceededPreflight ()
		{
			string text = CreateTempSystemRoot ();
			CreateWordPackage (Path.Combine (GetTemplateDirectory (text), "01_委任状.docx"), BuildDocumentXml (CreateTextControl ("顧客_名前")));
			CreateWordPackage (Path.Combine (GetTemplateDirectory (text), "05_照会書.dotx"), BuildDocumentXml ());

			KernelTemplateSyncPreflightResult kernelTemplateSyncPreflightResult = CreateService ().Run (new KernelTemplateSyncPreflightRequest (text, new string[1] { "顧客_名前" }));

			Assert.Equal (KernelTemplateSyncPreflightStatus.Succeeded, kernelTemplateSyncPreflightResult.Status);
			Assert.NotNull (kernelTemplateSyncPreflightResult.ValidationSummary);
			Assert.Null (kernelTemplateSyncPreflightResult.Failure);
			Assert.Equal (2, kernelTemplateSyncPreflightResult.ValidationSummary.DetectedFileCount);
			Assert.Equal (new string[2] { "01", "05" }, kernelTemplateSyncPreflightResult.ValidationSummary.GetValidTemplates ().Select (entry => entry.Key).ToArray ());
		}

		[Fact]
		public void Run_WhenDetectedTemplatesContainValidationErrors_KeepsSucceededPreflight ()
		{
			string text = CreateTempSystemRoot ();
			CreateWordPackage (Path.Combine (GetTemplateDirectory (text), "21_契約書.docx"), BuildDocumentXml (CreateTextControl ("相手方_名前")));

			KernelTemplateSyncPreflightResult kernelTemplateSyncPreflightResult = CreateService ().Run (new KernelTemplateSyncPreflightRequest (text, new string[1] { "顧客_名前" }));

			Assert.Equal (KernelTemplateSyncPreflightStatus.Succeeded, kernelTemplateSyncPreflightResult.Status);
			Assert.NotNull (kernelTemplateSyncPreflightResult.ValidationSummary);
			Assert.Null (kernelTemplateSyncPreflightResult.Failure);
			Assert.Single (kernelTemplateSyncPreflightResult.ValidationSummary.TemplateResults);
			Assert.Empty (kernelTemplateSyncPreflightResult.ValidationSummary.GetValidTemplates ());
			Assert.Contains ("未定義タグ「相手方_名前」があります。", kernelTemplateSyncPreflightResult.ValidationSummary.TemplateResults[0].Errors);
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

		private static KernelTemplateSyncPreflightService CreateService ()
		{
			Logger logger = OrchestrationTestSupport.CreateLogger (new List<string> ());
			return new KernelTemplateSyncPreflightService (new PathCompatibilityService (logger), new WordTemplateRegistrationValidationService (new WordTemplateContentControlInspectionService (), logger));
		}

		private string CreateTempSystemRoot ()
		{
			string path = Path.Combine (Path.GetTempPath (), "CaseInfoSystem.KernelTemplateSyncPreflight." + Guid.NewGuid ().ToString ("N"));
			Directory.CreateDirectory (path);
			Directory.CreateDirectory (GetTemplateDirectory (path));
			_temporaryDirectories.Add (path);
			return path;
		}

		private static string GetTemplateDirectory (string systemRoot)
		{
			return Path.Combine (systemRoot, "雛形");
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

		private static string BuildTagElement (string tag)
		{
			return (tag == null) ? string.Empty : "<w:tag w:val=\"" + SecurityElement.Escape (tag) + "\"/>";
		}
	}
}
