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
		public void SaveDocument_WhenFinalFileExists_SavesDirectlyToFinalPathAfterCreatingAdjacentBackup ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			List<string> backupFilesDuringSave = new List<string> ();
			string backupContentDuringSave = string.Empty;
			using TestFolderScope scope = new TestFolderScope ();
			string preparedFinalPath = Path.Combine (scope.RootPath, "inside.docx");
			File.WriteAllText (preparedFinalPath, "old");

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					SaveDocumentAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						backupFilesDuringSave.AddRange (Directory.GetFiles (scope.RootPath, "inside.bak_*" + Path.GetExtension (preparedFinalPath)));
						backupContentDuringSave = File.ReadAllText (backupFilesDuringSave[0]);
						File.WriteAllText (savePath, "new");
						return savePath;
					}
				});

			DocumentSaveResult result = context.Service.SaveDocument (new object (), new object (), preparedFinalPath);

			Assert.Equal (new[] { preparedFinalPath }, savedPaths);
			Assert.Single (backupFilesDuringSave);
			Assert.Equal ("old", backupContentDuringSave);
			Assert.Equal ("new", File.ReadAllText (preparedFinalPath));
			Assert.Empty (Directory.GetFiles (scope.RootPath, "inside.bak_*" + Path.GetExtension (preparedFinalPath)));
			Assert.Contains (logs, log => log.Contains ("SaveContext mode=Overwrite location=NonSyncRoot"));
			Assert.False (result.IsLocalWorkCopy);
			Assert.Equal (preparedFinalPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
		}

		[Fact]
		public void SaveDocument_WhenFinalFileDoesNotExist_SavesDirectlyWithoutCreatingBackup ()
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
					SaveDocumentAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						Assert.Empty (Directory.GetFiles (scope.RootPath, "new.bak_*" + Path.GetExtension (preparedFinalPath)));
						File.WriteAllText (savePath, "new");
						return savePath;
					}
				});

			DocumentSaveResult result = context.Service.SaveDocument (new object (), new object (), preparedFinalPath);

			Assert.Equal (new[] { preparedFinalPath }, savedPaths);
			Assert.Equal ("new", File.ReadAllText (preparedFinalPath));
			Assert.Empty (Directory.GetFiles (scope.RootPath, "new.bak_*" + Path.GetExtension (preparedFinalPath)));
			Assert.Contains (logs, log => log.Contains ("SaveContext mode=CreateNew location=NonSyncRoot"));
			Assert.False (result.IsLocalWorkCopy);
			Assert.Equal (preparedFinalPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
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
						SaveDocumentAsDocx = (document, savePath) =>
						{
							File.WriteAllText (savePath, "new");
							return savePath;
						}
					});

				context.Service.SaveDocument (new object (), new object (), preparedFinalPath);
			} finally {
				Environment.SetEnvironmentVariable ("OneDrive", originalOneDrive);
			}

			Assert.Contains (logs, log => log.Contains ("SaveContext mode=CreateNew location=SyncRoot"));
		}

		[Fact]
		public void SaveDocument_WhenDirectSaveFails_RestoresExistingFinalFileFromBackup ()
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
					SaveDocumentAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						File.WriteAllText (savePath, "partial");
						throw new IOException ("save failed");
					}
				});

			IOException exception = Assert.Throws<IOException> (() => context.Service.SaveDocument (new object (), new object (), preparedFinalPath));

			Assert.Equal ("save failed", exception.Message);
			Assert.Equal (new[] { preparedFinalPath }, savedPaths);
			Assert.Equal ("old", File.ReadAllText (preparedFinalPath));
			Assert.Empty (Directory.GetFiles (scope.RootPath, "restore.bak_*" + Path.GetExtension (preparedFinalPath)));
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
	}
}
