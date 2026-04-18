using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class LocalWorkCopyServiceTests
	{
		[Fact]
		public void RegisterLocalWorkCopy_WhenPathsAreValid_SetsPendingStateAndStartsPolling ()
		{
			List<string> logs = new List<string> ();

			using TestServiceContext context = CreateContext (logs);

			context.Service.RegisterLocalWorkCopy (new object (), @" C:/temp/draft.docx ", @" C:\final\result.docx ");

			Assert.True (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal ("- result.docx", context.Service.GetPendingLocalWorkCopySummary ());
			Assert.True (context.Service.IsPollingActiveForTesting ());
			Assert.Contains (logs, message => message.Contains ("registered local copy"));
		}

		[Fact]
		public void RegisterLocalWorkCopy_WhenSameLocalPathAlreadyTracked_ReplacesExistingRegistration ()
		{
			using TestServiceContext context = CreateContext (new List<string> ());

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\old.docx");
			context.Service.RegisterLocalWorkCopy (new object (), @"c:/temp/draft.docx", @"C:\final\new.docx");

			string summary = context.Service.GetPendingLocalWorkCopySummary ();

			Assert.True (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal ("- new.docx", summary);
			Assert.DoesNotContain ("old.docx", summary);
			Assert.Single (SplitSummaryLines (summary));
		}

		[Fact]
		public void Cancel_WhenPendingRegistrationExists_StopsPollingAndKeepsPendingState ()
		{
			using TestServiceContext context = CreateContext (new List<string> ());

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\result.docx");
			context.Service.Cancel ();

			Assert.True (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal ("- result.docx", context.Service.GetPendingLocalWorkCopySummary ());
			Assert.False (context.Service.IsPollingActiveForTesting ());
		}

		[Fact]
		public void ExecutePollLocalWorkCopiesForTesting_WhenDocumentIsStillOpen_KeepsPendingState ()
		{
			bool moveCalled = false;

			using TestServiceContext context = CreateContext (
				new List<string> (),
				new LocalWorkCopyService.LocalWorkCopyServiceTestHooks
				{
					IsDocumentOpen = (application, path) => true,
					MoveFileSafe = (source, destination) =>
					{
						moveCalled = true;
						return true;
					}
				});

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\result.docx");
			context.Service.ExecutePollLocalWorkCopiesForTesting ();

			Assert.True (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal ("- result.docx", context.Service.GetPendingLocalWorkCopySummary ());
			Assert.True (context.Service.IsPollingActiveForTesting ());
			Assert.False (moveCalled);
		}

		[Fact]
		public void ExecutePollLocalWorkCopiesForTesting_WhenDocumentClosedAndMoveSucceeds_FinalizesAndStopsPolling ()
		{
			string movedSource = null;
			string movedDestination = null;
			List<string> logs = new List<string> ();

			using TestServiceContext context = CreateContext (
				logs,
				new LocalWorkCopyService.LocalWorkCopyServiceTestHooks
				{
					IsDocumentOpen = (application, path) => false,
					FileExistsSafe = path => true,
					MoveFileSafe = (source, destination) =>
					{
						movedSource = source;
						movedDestination = destination;
						return true;
					}
				});

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\result.docx");
			context.Service.ExecutePollLocalWorkCopiesForTesting ();

			Assert.False (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal (string.Empty, context.Service.GetPendingLocalWorkCopySummary ());
			Assert.False (context.Service.IsPollingActiveForTesting ());
			Assert.Equal (@"C:\temp\draft.docx", movedSource);
			Assert.Equal (@"C:\final\result.docx", movedDestination);
			Assert.Contains (logs, message => message.Contains ("finalized local copy"));
		}

		[Fact]
		public void ExecutePollLocalWorkCopiesForTesting_WhenLocalCopyIsMissing_RemovesPendingJobWithoutMove ()
		{
			bool moveCalled = false;

			using TestServiceContext context = CreateContext (
				new List<string> (),
				new LocalWorkCopyService.LocalWorkCopyServiceTestHooks
				{
					IsDocumentOpen = (application, path) => false,
					FileExistsSafe = path => false,
					MoveFileSafe = (source, destination) =>
					{
						moveCalled = true;
						return true;
					}
				});

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\result.docx");
			context.Service.ExecutePollLocalWorkCopiesForTesting ();

			Assert.False (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal (string.Empty, context.Service.GetPendingLocalWorkCopySummary ());
			Assert.False (context.Service.IsPollingActiveForTesting ());
			Assert.False (moveCalled);
		}

		[Fact]
		public void ExecutePollLocalWorkCopiesForTesting_WhenMoveFails_KeepsPendingStateAndShowsWarning ()
		{
			int warningCount = 0;
			string warningLocalPath = null;
			string warningFinalPath = null;

			using TestServiceContext context = CreateContext (
				new List<string> (),
				new LocalWorkCopyService.LocalWorkCopyServiceTestHooks
				{
					IsDocumentOpen = (application, path) => false,
					FileExistsSafe = path => true,
					MoveFileSafe = (source, destination) => false,
					ShowFinalizeFailureMessage = (localPath, finalPath) =>
					{
						warningCount++;
						warningLocalPath = localPath;
						warningFinalPath = finalPath;
					}
				});

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\result.docx");
			context.Service.ExecutePollLocalWorkCopiesForTesting ();

			Assert.True (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal ("- result.docx", context.Service.GetPendingLocalWorkCopySummary ());
			Assert.True (context.Service.IsPollingActiveForTesting ());
			Assert.Equal (1, warningCount);
			Assert.Equal (@"C:\temp\draft.docx", warningLocalPath);
			Assert.Equal (@"C:\final\result.docx", warningFinalPath);
		}

		[Fact]
		public void RaisePollTimerTickForTesting_WhenPollingThrows_LogsErrorAndKeepsPendingState ()
		{
			List<string> logs = new List<string> ();

			using TestServiceContext context = CreateContext (
				logs,
				new LocalWorkCopyService.LocalWorkCopyServiceTestHooks
				{
					IsDocumentOpen = (application, path) => throw new InvalidOperationException ("boom")
				});

			context.Service.RegisterLocalWorkCopy (new object (), @"C:\temp\draft.docx", @"C:\final\result.docx");
			context.Service.RaisePollTimerTickForTesting ();

			Assert.True (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal ("- result.docx", context.Service.GetPendingLocalWorkCopySummary ());
			Assert.True (context.Service.IsPollingActiveForTesting ());
			Assert.Contains (logs, message => message.Contains ("LocalWorkCopyService.PollLocalWorkCopies failed."));
			Assert.Contains (logs, message => message.Contains ("InvalidOperationException: boom"));
		}

		[Fact]
		public void RegisterLocalWorkCopy_WhenNormalizedPathIsEmpty_DoesNotChangeState ()
		{
			using TestServiceContext context = CreateContext (new List<string> ());

			context.Service.RegisterLocalWorkCopy (new object (), " ", @"C:\final\result.docx");

			Assert.False (context.Service.HasPendingLocalWorkCopies ());
			Assert.Equal (string.Empty, context.Service.GetPendingLocalWorkCopySummary ());
			Assert.False (context.Service.IsPollingActiveForTesting ());
		}

		private static string[] SplitSummaryLines (string summary)
		{
			return (summary ?? string.Empty).Split (new[] { Environment.NewLine }, StringSplitOptions.None);
		}

		private static TestServiceContext CreateContext (List<string> logs, LocalWorkCopyService.LocalWorkCopyServiceTestHooks testHooks = null)
		{
			Logger logger = OrchestrationTestSupport.CreateLogger (logs);
			var pathCompatibilityService = new PathCompatibilityService ();
			var wordInteropService = new WordInteropService (pathCompatibilityService, logger);
			var service = new LocalWorkCopyService (pathCompatibilityService, wordInteropService, logger, testHooks);
			return new TestServiceContext (service);
		}

		private sealed class TestServiceContext : IDisposable
		{
			internal TestServiceContext (LocalWorkCopyService service)
			{
				Service = service;
			}

			internal LocalWorkCopyService Service { get; }

			public void Dispose ()
			{
				Service.Dispose ();
			}
		}
	}
}
