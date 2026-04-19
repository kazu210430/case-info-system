using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class WordInteropServiceTests
	{
		[Fact]
		public void AcquireWordApplication_WhenCachedInstanceIsAlive_ReusesCachedInstanceWithoutCreatingNew ()
		{
			List<string> logs = new List<string> ();
			object existingWord = new object ();
			int getActiveObjectCalls = 0;
			int createInstanceCalls = 0;
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				IsWordApplicationAlive = application => ReferenceEquals (application, existingWord),
				GetActiveObject = progId =>
				{
					getActiveObjectCalls++;
					return existingWord;
				},
				GetTypeFromProgID = progId => typeof (object),
				CreateInstance = type =>
				{
					createInstanceCalls++;
					return new object ();
				}
			};
			WordInteropService service = CreateService (logs, testHooks);

			object first = service.AcquireWordApplication (out bool createdNewFirst);
			object second = service.AcquireWordApplication (out bool createdNewSecond);

			Assert.Same (existingWord, first);
			Assert.Same (existingWord, second);
			Assert.False (createdNewFirst);
			Assert.False (createdNewSecond);
			Assert.Equal (1, getActiveObjectCalls);
			Assert.Equal (0, createInstanceCalls);
			Assert.Empty (logs);
		}

		[Fact]
		public void AcquireWordApplication_WhenReuseTargetIsNotAlive_FallsBackToCreatingNewInstance ()
		{
			object staleWord = new object ();
			object newWord = new object ();
			int createInstanceCalls = 0;
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				IsWordApplicationAlive = application => ReferenceEquals (application, newWord),
				GetActiveObject = progId => staleWord,
				GetTypeFromProgID = progId => typeof (object),
				CreateInstance = type =>
				{
					createInstanceCalls++;
					return newWord;
				}
			};
			WordInteropService service = CreateService (new List<string> (), testHooks);

			object result = service.AcquireWordApplication (out bool createdNew);

			Assert.Same (newWord, result);
			Assert.True (createdNew);
			Assert.Equal (1, createInstanceCalls);
		}

		[Fact]
		public void EnsureWordApplication_WhenCurrentInstanceIsAlive_ReturnsSameInstanceWithoutReacquiring ()
		{
			object existingWord = new object ();
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				IsWordApplicationAlive = application => ReferenceEquals (application, existingWord),
				GetActiveObject = progId => throw new InvalidOperationException ("should not reacquire"),
				GetTypeFromProgID = progId => throw new InvalidOperationException ("should not resolve type")
			};
			WordInteropService service = CreateService (new List<string> (), testHooks);
			object wordApplication = existingWord;

			object result = service.EnsureWordApplication (ref wordApplication);

			Assert.Same (existingWord, result);
			Assert.Same (existingWord, wordApplication);
		}

		[Fact]
		public void EnsureWordApplication_WhenCurrentInstanceIsDead_ReplacesReferenceWithAcquiredInstance ()
		{
			object staleWord = new object ();
			object acquiredWord = new object ();
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				IsWordApplicationAlive = application => ReferenceEquals (application, acquiredWord),
				GetActiveObject = progId => acquiredWord
			};
			WordInteropService service = CreateService (new List<string> (), testHooks);
			object wordApplication = staleWord;

			object result = service.EnsureWordApplication (ref wordApplication);

			Assert.Same (acquiredWord, result);
			Assert.Same (acquiredWord, wordApplication);
		}

		[Fact]
		public void AcquireWordApplication_WhenCreateInstanceThrowsComException_ReturnsNullAndLogsError ()
		{
			List<string> logs = new List<string> ();
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				IsWordApplicationAlive = application => false,
				GetActiveObject = progId => null,
				GetTypeFromProgID = progId => typeof (object),
				CreateInstance = type => throw new COMException ("create failed")
			};
			WordInteropService service = CreateService (logs, testHooks);

			object result = service.AcquireWordApplication (out bool createdNew);

			Assert.Null (result);
			Assert.False (createdNew);
			Assert.Contains (logs, message => message.Contains ("WordInteropService.AcquireWordApplication failed."));
			Assert.Contains (logs, message => message.Contains ("COMException: create failed"));
		}

		[Fact]
		public void CreateDocumentFromTemplate_WhenTemplatePathDoesNotExist_ReturnsNullAndDoesNotCallAdd ()
		{
			List<string> logs = new List<string> ();
			WordInteropService service = CreateService (logs, testHooks: null);
			var wordApplication = new TemplateWordApplication ();
			string templatePath = Path.Combine (Path.GetTempPath (), Guid.NewGuid ().ToString ("N"), "missing.dotx");

			object result = service.CreateDocumentFromTemplate (wordApplication, templatePath);

			Assert.Null (result);
			Assert.Equal (0, wordApplication.Documents.AddCalls);
			Assert.Contains (logs, message => message.Contains ("WordInteropService.CreateDocumentFromTemplate template not found."));
		}

		[Fact]
		public void CreateDocumentFromTemplate_WhenTemplatePathResolutionThrows_ReturnsNullAndLogsError ()
		{
			List<string> logs = new List<string> ();
			var wordApplication = new TemplateWordApplication ();
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				ResolveToExistingLocalPath = path => throw new InvalidOperationException ("resolve failed")
			};
			WordInteropService service = CreateService (logs, testHooks);

			object result = service.CreateDocumentFromTemplate (wordApplication, "template.dotx");

			Assert.Null (result);
			Assert.Equal (0, wordApplication.Documents.AddCalls);
			Assert.Contains (logs, message => message.Contains ("WordInteropService.CreateDocumentFromTemplate template path resolution failed."));
			Assert.Contains (logs, message => message.Contains ("InvalidOperationException: resolve failed"));
		}

		[Fact]
		public void CreateDocumentFromTemplate_WhenDocumentsAddThrowsComException_ReturnsNullAndLogsError ()
		{
			List<string> logs = new List<string> ();
			WordInteropService service = CreateService (logs, testHooks: null);
			var wordApplication = new ThrowingAddWordApplication ();
			string templatePath = Path.GetTempFileName ();
			try {
				object result = service.CreateDocumentFromTemplate (wordApplication, templatePath);

				Assert.Null (result);
				Assert.Equal (1, wordApplication.Documents.AddCalls);
				Assert.Contains (logs, message => message.Contains ("WordInteropService.CreateDocumentFromTemplate Documents.Add failed."));
				Assert.Contains (logs, message => message.Contains ("COMException: add failed"));
			} finally {
				File.Delete (templatePath);
			}
		}

		[Fact]
		public void CreateDocumentFromTemplate_WhenWordApplicationIsNull_ReturnsNullWithoutResolvingTemplate ()
		{
			List<string> logs = new List<string> ();
			int resolveCalls = 0;
			var testHooks = new WordInteropService.WordInteropServiceTestHooks
			{
				ResolveToExistingLocalPath = path =>
				{
					resolveCalls++;
					return path;
				}
			};
			WordInteropService service = CreateService (logs, testHooks);

			object result = service.CreateDocumentFromTemplate (wordApplication: null, templatePath: "template.dotx");

			Assert.Null (result);
			Assert.Equal (0, resolveCalls);
			Assert.Contains (logs, message => message.Contains ("WordInteropService.CreateDocumentFromTemplate skipped because wordApplication was null."));
		}

		[Fact]
		public void CreateDocumentFromTemplate_WhenTemplateExists_ReturnsCreatedDocument ()
		{
			List<string> logs = new List<string> ();
			WordInteropService service = CreateService (logs, testHooks: null);
			var wordApplication = new TemplateWordApplication ();
			object createdDocument = new object ();
			wordApplication.Documents.AddResult = createdDocument;
			string templatePath = Path.GetTempFileName ();
			try {
				object result = service.CreateDocumentFromTemplate (wordApplication, templatePath);

				Assert.Same (createdDocument, result);
				Assert.Equal (1, wordApplication.Documents.AddCalls);
				Assert.Equal (new PathCompatibilityService ().NormalizePath (templatePath), wordApplication.Documents.LastTemplatePath);
				Assert.False (wordApplication.Documents.LastNewTemplate);
				Assert.Empty (logs);
			} finally {
				File.Delete (templatePath);
			}
		}

		[Fact]
		public void IsDocumentOpen_WhenMatchingDocumentExists_ReturnsTrue ()
		{
			WordInteropService service = CreateService (new List<string> (), testHooks: null);
			var wordApplication = new FakeWordApplication ();
			wordApplication.Documents.Add (new FakeWordDocument { FullName = @"C:\Docs\Test.docx" });

			bool isOpen = service.IsDocumentOpen (wordApplication, @"c:/docs/test.docx");

			Assert.True (isOpen);
		}

		[Fact]
		public void IsDocumentOpen_WhenDocumentsEnumerationThrowsComException_ReturnsFalse ()
		{
			WordInteropService service = CreateService (new List<string> (), testHooks: null);

			bool isOpen = service.IsDocumentOpen (new ThrowingDocumentsWordApplication (), @"C:\Docs\Test.docx");

			Assert.False (isOpen);
		}

		private static WordInteropService CreateService (List<string> logs, WordInteropService.WordInteropServiceTestHooks testHooks)
		{
			Logger logger = OrchestrationTestSupport.CreateLogger (logs);
			var pathCompatibilityService = new PathCompatibilityService ();
			return new WordInteropService (pathCompatibilityService, logger, testHooks);
		}

		public sealed class FakeWordApplication
		{
			public List<object> Documents { get; } = new List<object> ();
		}

		public sealed class TemplateWordApplication
		{
			public RecordingDocumentsCollection Documents { get; } = new RecordingDocumentsCollection ();
		}

		public sealed class FakeWordDocument
		{
			public string FullName { get; set; }
		}

		public sealed class RecordingDocumentsCollection
		{
			public int AddCalls { get; private set; }

			public string LastTemplatePath { get; private set; }

			public bool LastNewTemplate { get; private set; }

			public object AddResult { get; set; } = new object ();

			public object Add (string templatePath, bool newTemplate)
			{
				AddCalls++;
				LastTemplatePath = templatePath;
				LastNewTemplate = newTemplate;
				return AddResult;
			}
		}

		public sealed class ThrowingDocumentsWordApplication
		{
			public IEnumerable<object> Documents => throw new COMException ("documents unavailable");
		}

		public sealed class ThrowingAddWordApplication
		{
			public ThrowingAddDocumentsCollection Documents { get; } = new ThrowingAddDocumentsCollection ();
		}

		public sealed class ThrowingAddDocumentsCollection
		{
			public int AddCalls { get; private set; }

			public object Add (string templatePath, bool newTemplate)
			{
				AddCalls++;
				throw new COMException ("add failed");
			}
		}
	}
}
