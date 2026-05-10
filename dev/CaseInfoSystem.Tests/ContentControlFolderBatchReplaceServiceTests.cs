using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using CaseInfoSystem.WordAddIn.Services;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class ContentControlFolderBatchReplaceServiceTests
	{
		private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		[Fact]
		public void Execute_ReplacesTextAndRichTextControlsOnlyInTopLevelTemplateFiles ()
		{
			string directory = CreateTempDirectory ();
			try {
				string documentPath = Path.Combine (directory, "01_通知書.docx");
				CreateWordPackage (
					documentPath,
					BuildDocumentXml (
						CreateTextControl ("OldTag", "OldTitle")
						+ CreateRichTextControl ("Case-001", "NoMatch")
						+ CreateCheckBoxControl ("OldTag", "OldTitle")));
				Directory.CreateDirectory (Path.Combine (directory, "Sub"));
				CreateWordPackage (
					Path.Combine (directory, "Sub", "02_下位.docx"),
					BuildDocumentXml (CreateTextControl ("OldTag", "OldTitle")));

				ContentControlFolderBatchReplaceService.FolderReplaceResult result = new ContentControlFolderBatchReplaceService ().Execute (
					new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
						TemplateDirectory = directory,
						OldTag = "Old",
						NewTag = "New",
						OldTitle = "OldTitle",
						NewTitle = "NewTitle",
						UsePartialMatch = true,
						CreateBackups = false
					});

				Assert.Equal (1, result.DetectedFileCount);
				Assert.Equal (1, result.ProcessedFileCount);
				Assert.Equal (1, result.ChangedFileCount);
				Assert.Equal (0, result.FailedFileCount);
				Assert.Equal (2, result.ScannedCount);
				Assert.Equal (1, result.TagChangedCount);
				Assert.Equal (1, result.TitleChangedCount);

				XElement[] controls = ReadContentControls (documentPath);
				Assert.Equal ("NewTag", ReadTag (controls[0]));
				Assert.Equal ("NewTitle", ReadTitle (controls[0]));
				Assert.Equal ("Case-001", ReadTag (controls[1]));
				Assert.Equal ("NoMatch", ReadTitle (controls[1]));
				Assert.Equal ("OldTag", ReadTag (controls[2]));
				Assert.Equal ("OldTitle", ReadTitle (controls[2]));

				XElement[] nestedControls = ReadContentControls (Path.Combine (directory, "Sub", "02_下位.docx"));
				Assert.Equal ("OldTag", ReadTag (nestedControls[0]));
				Assert.Equal ("OldTitle", ReadTitle (nestedControls[0]));
			} finally {
				DeleteDirectoryQuietly (directory);
			}
		}

		[Fact]
		public void Execute_WhenDisplayTextMatches_ReplacesContentControlDisplayText ()
		{
			string directory = CreateTempDirectory ();
			try {
				string documentPath = Path.Combine (directory, "01_通知書.docx");
				CreateWordPackage (
					documentPath,
					BuildDocumentXml (
						CreateTextControl ("TagA", "TitleA", "<w:r><w:t>旧</w:t></w:r><w:r><w:t>表示</w:t></w:r>")
						+ CreateCheckBoxControl ("TagB", "TitleB", "<w:r><w:t>旧表示</w:t></w:r>")));

				ContentControlFolderBatchReplaceService.FolderReplaceResult result = new ContentControlFolderBatchReplaceService ().Execute (
					new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
						TemplateDirectory = directory,
						OldDisplayText = "旧表示",
						NewDisplayText = "新表示",
						CreateBackups = false
					});

				Assert.Equal (1, result.ChangedFileCount);
				Assert.Equal (1, result.DisplayTextChangedCount);
				Assert.Equal (0, result.TagChangedCount);
				Assert.Equal (0, result.TitleChangedCount);

				XElement[] controls = ReadContentControls (documentPath);
				Assert.Equal ("新表示", ReadDisplayText (controls[0]));
				Assert.Equal ("旧表示", ReadDisplayText (controls[1]));
			} finally {
				DeleteDirectoryQuietly (directory);
			}
		}

		[Fact]
		public void Execute_WhenDisplayTextPartialMatch_ReplacesAcrossSplitRuns ()
		{
			string directory = CreateTempDirectory ();
			try {
				string documentPath = Path.Combine (directory, "01_通知書.docx");
				CreateWordPackage (
					documentPath,
					BuildDocumentXml (
						CreateTextControl ("TagA", "TitleA", "<w:r><w:t>当</w:t></w:r><w:r><w:t>方_郵便番号</w:t></w:r>")));

				ContentControlFolderBatchReplaceService.FolderReplaceResult result = new ContentControlFolderBatchReplaceService ().Execute (
					new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
						TemplateDirectory = directory,
						OldDisplayText = "当方_",
						NewDisplayText = "My_",
						UsePartialMatch = true,
						CreateBackups = false
					});

				Assert.Equal (1, result.DisplayTextChangedCount);
				Assert.Equal ("My_郵便番号", ReadDisplayText (ReadContentControls (documentPath)[0]));
			} finally {
				DeleteDirectoryQuietly (directory);
			}
		}

		[Fact]
		public void Execute_WhenDisplayTextHasLayoutSpaces_ReplacesCompactTextAndKeepsSpacing ()
		{
			string directory = CreateTempDirectory ();
			try {
				string documentPath = Path.Combine (directory, "01_Template.docx");
				CreateWordPackage (
					documentPath,
					BuildDocumentXml (
						CreateTextControl (
							"TagA",
							"TitleA",
							"<w:r><w:t>\u3010 M y \u5f01 \u8b77 \u58eb \u3011</w:t></w:r>")));

				ContentControlFolderBatchReplaceService.FolderReplaceResult result = new ContentControlFolderBatchReplaceService ().Execute (
					new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
						TemplateDirectory = directory,
						OldDisplayText = "My\u5f01\u8b77\u58eb",
						NewDisplayText = "My\u8868\u793a\u540d",
						UsePartialMatch = true,
						CreateBackups = false
					});

				Assert.Equal (1, result.DisplayTextChangedCount);
				Assert.Equal ("\u3010 M y \u8868 \u793a \u540d \u3011", ReadDisplayText (ReadContentControls (documentPath)[0]));
			} finally {
				DeleteDirectoryQuietly (directory);
			}
		}

		[Fact]
		public void Execute_WhenBackupEnabled_CreatesBackupBeforeWritingChanges ()
		{
			string directory = CreateTempDirectory ();
			try {
				string documentPath = Path.Combine (directory, "01_委任状.dotx");
				CreateWordPackage (documentPath, BuildDocumentXml (CreateTextControl ("OldTag", "OldTitle")));

				ContentControlFolderBatchReplaceService.FolderReplaceResult result = new ContentControlFolderBatchReplaceService ().Execute (
					new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
						TemplateDirectory = directory,
						OldTag = "OldTag",
						NewTag = "NewTag",
						CreateBackups = true
					});

				Assert.Equal (1, result.BackupCreatedCount);
				ContentControlFolderBatchReplaceService.FileReplaceResult fileResult = Assert.Single (result.FileResults);
				Assert.True (fileResult.BackupCreated);
				Assert.True (File.Exists (fileResult.BackupPath));
				Assert.Equal ("OldTag", ReadTag (ReadContentControls (fileResult.BackupPath)[0]));
				Assert.Equal ("NewTag", ReadTag (ReadContentControls (documentPath)[0]));
			} finally {
				DeleteDirectoryQuietly (directory);
			}
		}

		[Fact]
		public void Execute_WhenOneFileIsInvalid_RecordsFailureAndContinues ()
		{
			string directory = CreateTempDirectory ();
			try {
				string invalidPath = Path.Combine (directory, "01_破損.docx");
				File.WriteAllText (invalidPath, "not a zip package");
				string validPath = Path.Combine (directory, "02_通知書.docx");
				CreateWordPackage (validPath, BuildDocumentXml (CreateTextControl ("OldTag", "OldTitle")));

				ContentControlFolderBatchReplaceService.FolderReplaceResult result = new ContentControlFolderBatchReplaceService ().Execute (
					new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
						TemplateDirectory = directory,
						OldTag = "OldTag",
						NewTag = "NewTag",
						CreateBackups = false
					});

				Assert.Equal (2, result.DetectedFileCount);
				Assert.Equal (1, result.ProcessedFileCount);
				Assert.Equal (1, result.ChangedFileCount);
				Assert.Equal (1, result.FailedFileCount);
				Assert.Contains (result.FileResults, item => item.FileName == "01_破損.docx" && !item.Success);
				Assert.Equal ("NewTag", ReadTag (ReadContentControls (validPath)[0]));
			} finally {
				DeleteDirectoryQuietly (directory);
			}
		}

		private static string CreateTempDirectory ()
		{
			string path = Path.Combine (Path.GetTempPath (), "CaseInfoSystem_" + Guid.NewGuid ().ToString ("N"));
			Directory.CreateDirectory (path);
			return path;
		}

		private static void DeleteDirectoryQuietly (string directory)
		{
			try {
				if (!string.IsNullOrWhiteSpace (directory) && Directory.Exists (directory)) {
					Directory.Delete (directory, true);
				}
			} catch {
			}
		}

		private static void CreateWordPackage (string fullPath, string documentXml)
		{
			using (ZipArchive zipArchive = ZipFile.Open (fullPath, ZipArchiveMode.Create)) {
				ZipArchiveEntry zipArchiveEntry = zipArchive.CreateEntry ("word/document.xml");
				using (Stream stream = zipArchiveEntry.Open ())
				using (StreamWriter streamWriter = new StreamWriter (stream)) {
					streamWriter.Write (documentXml);
				}
			}
		}

		private static XElement[] ReadContentControls (string fullPath)
		{
			using (ZipArchive zipArchive = ZipFile.OpenRead (fullPath)) {
				ZipArchiveEntry zipArchiveEntry = zipArchive.GetEntry ("word/document.xml");
				using (Stream stream = zipArchiveEntry.Open ()) {
					return XDocument.Load (stream).Descendants (W + "sdt").ToArray ();
				}
			}
		}

		private static string ReadTag (XElement contentControl)
		{
			return ReadVal (contentControl, "tag");
		}

		private static string ReadTitle (XElement contentControl)
		{
			return ReadVal (contentControl, "alias");
		}

		private static string ReadDisplayText (XElement contentControl)
		{
			XElement content = contentControl.Element (W + "sdtContent");
			return content == null ? string.Empty : string.Concat (content.Descendants (W + "t").Select (item => item.Value));
		}

		private static string ReadVal (XElement contentControl, string elementName)
		{
			XElement properties = contentControl.Element (W + "sdtPr");
			XElement element = properties == null ? null : properties.Element (W + elementName);
			XAttribute attribute = element == null ? null : element.Attribute (W + "val");
			return attribute == null ? string.Empty : attribute.Value;
		}

		private static string BuildDocumentXml (string body)
		{
			return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
				+ "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
				+ "<w:body>"
				+ body
				+ "</w:body>"
				+ "</w:document>";
		}

		private static string CreateTextControl (string tag, string title)
		{
			return CreateControl (tag, title, "<w:text/>");
		}

		private static string CreateTextControl (string tag, string title, string contentXml)
		{
			return CreateControl (tag, title, "<w:text/>", contentXml);
		}

		private static string CreateRichTextControl (string tag, string title)
		{
			return CreateControl (tag, title, "<w:richText/>");
		}

		private static string CreateCheckBoxControl (string tag, string title)
		{
			return CreateControl (tag, title, "<w14:checkbox xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"/>");
		}

		private static string CreateCheckBoxControl (string tag, string title, string contentXml)
		{
			return CreateControl (tag, title, "<w14:checkbox xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"/>", contentXml);
		}

		private static string CreateControl (string tag, string title, string typeElement)
		{
			return CreateControl (tag, title, typeElement, "<w:p><w:r><w:t>sample</w:t></w:r></w:p>");
		}

		private static string CreateControl (string tag, string title, string typeElement, string contentXml)
		{
			return "<w:sdt><w:sdtPr>"
				+ "<w:alias w:val=\"" + title + "\"/>"
				+ "<w:tag w:val=\"" + tag + "\"/>"
				+ typeElement
				+ "</w:sdtPr><w:sdtContent>"
				+ contentXml
				+ "</w:sdtContent></w:sdt>";
		}
	}
}
