using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class DocumentSaveServiceTests
	{
		[Fact]
		public void SaveDocument_WhenPathIsOutsideSyncRoot_SavesDirectly ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			bool buildLocalWorkCopyPathCalled = false;
			bool registerLocalWorkCopyCalled = false;
			const string requestedFinalPath = @"C:\requested\outside.docx";
			const string preparedFinalPath = @"C:\resolved\outside.docx";

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					IsUnderSyncRoot = _ => false,
					BuildLocalWorkCopyPath = _ =>
					{
						buildLocalWorkCopyPathCalled = true;
						return @"C:\temp\unused.docx";
					},
					SaveDocumentAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						return savePath;
					},
					RegisterLocalWorkCopy = (application, localPath, finalPath) => registerLocalWorkCopyCalled = true
				});

			DocumentSaveResult result = context.Service.SaveDocument (new object (), new object (), requestedFinalPath);

			Assert.Equal (new[] { preparedFinalPath }, savedPaths);
			Assert.False (buildLocalWorkCopyPathCalled);
			Assert.False (registerLocalWorkCopyCalled);
			Assert.False (result.IsLocalWorkCopy);
			Assert.Equal (preparedFinalPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
			Assert.DoesNotContain (logs, message => message.Contains ("fell back to direct save"));
			Assert.DoesNotContain (logs, message => message.Contains ("saved via local work copy"));
		}

		[Fact]
		public void SaveDocument_WhenLocalWorkCopyPathExists_SavesViaLocalWorkCopy ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			object registeredWordApplication = null;
			string registeredLocalPath = null;
			string registeredFinalPath = null;
			const string preparedFinalPath = @"C:\Users\kazu2\OneDrive\Docs\inside.docx";
			const string localWorkCopyPath = @"C:\Users\kazu2\AppData\Local\CaseDocTemp\inside.docx";
			object wordApplication = new object ();

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					IsUnderSyncRoot = _ => true,
					BuildLocalWorkCopyPath = _ => localWorkCopyPath,
					SaveDocumentAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						return savePath;
					},
					RegisterLocalWorkCopy = (application, localPath, finalPath) =>
					{
						registeredWordApplication = application;
						registeredLocalPath = localPath;
						registeredFinalPath = finalPath;
					}
				});

			DocumentSaveResult result = context.Service.SaveDocument (wordApplication, new object (), preparedFinalPath);

			Assert.Equal (new[] { localWorkCopyPath }, savedPaths);
			Assert.Same (wordApplication, registeredWordApplication);
			Assert.Equal (localWorkCopyPath, registeredLocalPath);
			Assert.Equal (preparedFinalPath, registeredFinalPath);
			Assert.True (result.IsLocalWorkCopy);
			Assert.Equal (localWorkCopyPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
			Assert.Contains (logs, message => message.Contains ("saved via local work copy"));
		}

		[Fact]
		public void SaveDocument_WhenLocalWorkCopyPathIsMissing_FallsBackToDirectSave ()
		{
			List<string> logs = new List<string> ();
			List<string> savedPaths = new List<string> ();
			bool registerLocalWorkCopyCalled = false;
			const string preparedFinalPath = @"C:\Users\kazu2\OneDrive\Docs\fallback.docx";

			using TestServiceContext context = CreateContext (
				logs,
				new DocumentSaveService.DocumentSaveServiceTestHooks
				{
					PrepareSavePath = _ => preparedFinalPath,
					IsUnderSyncRoot = _ => true,
					BuildLocalWorkCopyPath = _ => string.Empty,
					SaveDocumentAsDocx = (document, savePath) =>
					{
						savedPaths.Add (savePath);
						return savePath;
					},
					RegisterLocalWorkCopy = (application, localPath, finalPath) => registerLocalWorkCopyCalled = true
				});

			DocumentSaveResult result = context.Service.SaveDocument (new object (), new object (), preparedFinalPath);

			Assert.Equal (new[] { preparedFinalPath }, savedPaths);
			Assert.False (registerLocalWorkCopyCalled);
			Assert.False (result.IsLocalWorkCopy);
			Assert.Equal (preparedFinalPath, result.SavedPath);
			Assert.Equal (preparedFinalPath, result.FinalPath);
			Assert.Contains (logs, message => message.Contains ("fell back to direct save"));
		}

		private static TestServiceContext CreateContext (List<string> logs, DocumentSaveService.DocumentSaveServiceTestHooks testHooks)
		{
			Logger logger = OrchestrationTestSupport.CreateLogger (logs);
			var pathCompatibilityService = new PathCompatibilityService ();
			var wordInteropService = new WordInteropService (pathCompatibilityService, logger);
			var localWorkCopyService = new LocalWorkCopyService (pathCompatibilityService, wordInteropService, logger);
			var documentOutputService = new DocumentOutputService (new ExcelInteropService (), pathCompatibilityService, logger);
			var service = new DocumentSaveService (documentOutputService, pathCompatibilityService, localWorkCopyService, wordInteropService, logger, testHooks);
			return new TestServiceContext (service, localWorkCopyService);
		}

		private sealed class TestServiceContext : System.IDisposable
		{
			internal TestServiceContext (DocumentSaveService service, LocalWorkCopyService localWorkCopyService)
			{
				Service = service;
				_localWorkCopyService = localWorkCopyService;
			}

			internal DocumentSaveService Service { get; }

			private readonly LocalWorkCopyService _localWorkCopyService;

			public void Dispose ()
			{
				_localWorkCopyService.Dispose ();
			}
		}
	}
}
