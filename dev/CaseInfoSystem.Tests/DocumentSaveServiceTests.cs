using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class DocumentSaveServiceTests
	{
		[Fact]
		public void SaveDocument_WhenFinalFileExists_SavesToAdjacentTempThenReplacesFinalFile ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			List<string> openedPaths = new List<string> ();
			object reopenedDocument = new object ();
			using TestFolderScope scope = new TestFolderScope ();
			string preparedFinalPath = Path.Combine (scope.RootPath, "inside.docx");
			File.WriteAllText (preparedFinalPath, "old");

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					SaveDocumentCopyAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						Assert.NotEqual (preparedFinalPath, savePath);
						Assert.Equal ("old", File.ReadAllText (preparedFinalPath));
						Assert.Equal (scope.RootPath, Path.GetDirectoryName (savePath));
						File.WriteAllText (savePath, "new");
						return savePath;
					},
					OpenDocument = (application, fullPath) =>
					{
						openedPaths.Add (fullPath);
						return reopenedDocument;
					}
				});

			DocumentSaveResult result = context.Service.SaveDocument (new object (), new ClosableWordDocument (), preparedFinalPath);

			Assert.Single (savedPaths);
			Assert.Single (openedPaths);
			Assert.Equal (preparedFinalPath, openedPaths[0]);
			Assert.Equal ("new", File.ReadAllText (preparedFinalPath));
			Assert.Empty (Directory.GetFiles (scope.RootPath, "inside.tmp_*" + Path.GetExtension (preparedFinalPath)));
			Assert.Contains (logs, log => log.Contains ("SaveContext mode=Overwrite location=NonSyncRoot"));
			Assert.False (result.IsLocalWorkCopy);
			Assert.Equal (preparedFinalPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
			Assert.Same (reopenedDocument, result.ActiveDocument);
		}

		[Fact]
		public void SaveDocument_WhenFinalFileDoesNotExist_StillUsesAdjacentTempBeforeFinalPlacement ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			using TestFolderScope scope = new TestFolderScope ();
			string preparedFinalPath = Path.Combine (scope.RootPath, "new.docx");

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					SaveDocumentCopyAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						Assert.NotEqual (preparedFinalPath, savePath);
						Assert.False (File.Exists (preparedFinalPath));
						File.WriteAllText (savePath, "new");
						return savePath;
					},
					OpenDocument = (application, fullPath) => new object ()
				});

			DocumentSaveResult result = context.Service.SaveDocument (new object (), new ClosableWordDocument (), preparedFinalPath);

			Assert.Single (savedPaths);
			Assert.Equal ("new", File.ReadAllText (preparedFinalPath));
			Assert.Empty (Directory.GetFiles (scope.RootPath, "new.tmp_*" + Path.GetExtension (preparedFinalPath)));
			Assert.Contains (logs, log => log.Contains ("SaveContext mode=CreateNew location=NonSyncRoot"));
			Assert.False (result.IsLocalWorkCopy);
			Assert.Equal (preparedFinalPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
			Assert.NotNull (result.ActiveDocument);
		}

		[Fact]
		public void SaveDocument_WhenFinalPathIsUnderSyncRoot_LogsSyncRootSaveContext ()
		{
			List<string> logs = new List<string> ();
			string originalOneDrive = Environment.GetEnvironmentVariable ("OneDrive");
			using TestFolderScope scope = new TestFolderScope ();
			string preparedFinalPath = Path.Combine (scope.RootPath, "sync.docx");

			try {
				Environment.SetEnvironmentVariable ("OneDrive", scope.RootPath);
				using TestServiceContext context = CreateContext (
					logs,
					new DocumentSaveService.DocumentSaveServiceTestHooks
					{
						PrepareSavePath = _ => preparedFinalPath,
						SaveDocumentCopyAsDocx = (document, savePath) =>
						{
							File.WriteAllText (savePath, "new");
							return savePath;
						},
						OpenDocument = (application, fullPath) => new object ()
					});

				context.Service.SaveDocument (new object (), new ClosableWordDocument (), preparedFinalPath);
			} finally {
				Environment.SetEnvironmentVariable ("OneDrive", originalOneDrive);
			}

			Assert.Contains (logs, log => log.Contains ("SaveContext mode=CreateNew location=SyncRoot"));
		}

		[Fact]
		public void SaveDocument_WhenStagingSaveFails_LeavesExistingFinalFileUntouched ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			using TestFolderScope scope = new TestFolderScope ();
			string preparedFinalPath = Path.Combine (scope.RootPath, "restore.docx");
			File.WriteAllText (preparedFinalPath, "old");

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					SaveDocumentCopyAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						File.WriteAllText (savePath, "partial");
						throw new IOException ("save failed");
					}
				});

			IOException exception = Assert.Throws<IOException> (() => context.Service.SaveDocument (new object (), new ClosableWordDocument (), preparedFinalPath));

			Assert.Equal ("save failed", exception.Message);
			Assert.Single (savedPaths);
			Assert.NotEqual (preparedFinalPath, savedPaths[0]);
			Assert.Equal ("old", File.ReadAllText (preparedFinalPath));
			Assert.Empty (Directory.GetFiles (scope.RootPath, "restore.tmp_*" + Path.GetExtension (preparedFinalPath)));
		}

		private static TestServiceContext CreateContext (List<string> logs, DocumentSaveService.DocumentSaveServiceTestHooks testHooks)
		{
			Logger logger = OrchestrationTestSupport.CreateLogger (logs);
			var pathCompatibilityService = new PathCompatibilityService ();
			var wordInteropService = new WordInteropService (pathCompatibilityService, logger);
			var documentOutputService = new DocumentOutputService (new ExcelInteropService (), pathCompatibilityService, logger);
			var service = new DocumentSaveService (documentOutputService, wordInteropService, logger, testHooks);
			return new TestServiceContext (service);
		}

		private sealed class TestServiceContext : IDisposable
		{
			internal TestServiceContext (DocumentSaveService service)
			{
				Service = service;
			}

			internal DocumentSaveService Service { get; }

			public void Dispose ()
			{
			}
		}

		private sealed class TestFolderScope : IDisposable
		{
			internal TestFolderScope ()
			{
				RootPath = Path.Combine (Path.GetTempPath (), "CaseInfoSystem.DocumentSaveServiceTests." + Guid.NewGuid ().ToString ("N"));
				Directory.CreateDirectory (RootPath);
			}

			internal string RootPath { get; }

			public void Dispose ()
			{
				try {
					if (Directory.Exists (RootPath)) {
						Directory.Delete (RootPath, recursive: true);
					}
				} catch {
				}
			}
		}

		private sealed class ClosableWordDocument
		{
			public bool Closed { get; private set; }

			public void Close (bool saveChanges)
			{
				Closed = true;
			}
		}
	}
}
