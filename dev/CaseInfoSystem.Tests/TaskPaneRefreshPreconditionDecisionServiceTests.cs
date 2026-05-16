using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRefreshPreconditionDecisionServiceTests
    {
        [Fact]
        public void Decide_ReturnsContinue_WhenNoPreconditionBlocks()
        {
            var service = new TaskPaneRefreshPreconditionDecisionService();
            Excel.Workbook workbook = new Excel.Workbook();
            Excel.Window window = new Excel.Window { Hwnd = 101 };

            TaskPaneRefreshPreconditionDecisionResult result = service.Decide(
                "WorkbookActivate",
                workbook,
                window,
                shouldIgnoreDuringProtection: () => false);

            Assert.True(result.CanRefresh);
            Assert.False(result.ShouldFailClosed);
            Assert.Equal(string.Empty, result.SkipReason);
            Assert.Equal(string.Empty, result.SkipActionName);
            Assert.False(result.FailClosedOutcome.HasValue);
            Assert.Null(result.NormalizedOutcomeAttemptResult);
            Assert.Equal(string.Empty, result.NormalizedOutcomeActionName);
        }

        [Fact]
        public void Decide_ReturnsFailClosedSkip_WhenWorkbookOpenWindowIsUnresolved()
        {
            var service = new TaskPaneRefreshPreconditionDecisionService();
            bool protectionProbeCalled = false;
            Excel.Workbook workbook = new Excel.Workbook();

            TaskPaneRefreshPreconditionDecisionResult result = service.Decide(
                "WorkbookOpen",
                workbook,
                window: null,
                shouldIgnoreDuringProtection: () =>
                {
                    protectionProbeCalled = true;
                    return true;
                });

            Assert.False(result.CanRefresh);
            Assert.True(result.ShouldFailClosed);
            Assert.False(protectionProbeCalled);
            Assert.True(result.FailClosedOutcome.HasValue);
            Assert.Equal("skip-workbook-open-window-dependent-refresh", result.SkipReason);
            Assert.Equal(result.SkipReason, result.SkipActionName);
            Assert.Equal(result.SkipActionName, result.NormalizedOutcomeActionName);
            Assert.Same(result.FailClosedOutcome.Value.AttemptResult, result.NormalizedOutcomeAttemptResult);
            Assert.False(result.NormalizedOutcomeAttemptResult.IsRefreshSucceeded);
            Assert.True(result.NormalizedOutcomeAttemptResult.WasSkipped);
            Assert.Equal(result.SkipActionName, result.NormalizedOutcomeAttemptResult.CompletionBasis);
        }

        [Fact]
        public void Decide_ReturnsFailClosedSkip_WhenProtectionBlocksRefresh()
        {
            var service = new TaskPaneRefreshPreconditionDecisionService();
            Excel.Workbook workbook = new Excel.Workbook();
            Excel.Window window = new Excel.Window { Hwnd = 202 };

            TaskPaneRefreshPreconditionDecisionResult result = service.Decide(
                "WindowActivate",
                workbook,
                window,
                shouldIgnoreDuringProtection: () => true);

            Assert.False(result.CanRefresh);
            Assert.True(result.ShouldFailClosed);
            Assert.True(result.FailClosedOutcome.HasValue);
            Assert.Equal("ignore-during-protection", result.SkipReason);
            Assert.Equal(result.SkipReason, result.SkipActionName);
            Assert.Equal(result.SkipActionName, result.NormalizedOutcomeActionName);
            Assert.Same(result.FailClosedOutcome.Value.AttemptResult, result.NormalizedOutcomeAttemptResult);
            Assert.True(result.NormalizedOutcomeAttemptResult.WasSkipped);
            Assert.Equal(
                ForegroundGuaranteeOutcomeStatus.SkippedNoKnownTarget,
                result.NormalizedOutcomeAttemptResult.ForegroundGuaranteeOutcome.Status);
        }

        [Fact]
        public void Source_DoesNotOwnTimerCallbackSessionCompletionOrDisplayExecution()
        {
            string source = ReadAppSource("TaskPaneRefreshPreconditionDecisionService.cs");

            Assert.Contains("TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(", source);
            Assert.Contains("TaskPaneRefreshFailClosedOutcome.FromPreconditionDecision(preconditionDecision)", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("TaskPaneReadyShowRetryScheduler", source);
            Assert.DoesNotContain("PendingPaneRefreshRetryService", source);
            Assert.DoesNotContain("WindowActivateDownstreamObservation", source);
            Assert.DoesNotContain("CompleteNormalizedOutcomeChain", source);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome", source);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", source);
            Assert.DoesNotContain("case-display-completed", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("TaskPaneRefreshCoordinator", source);
            Assert.DoesNotContain("TaskPaneManager", source);
            Assert.DoesNotContain(".Hwnd", source);
            Assert.DoesNotContain(".FullName", source);
            Assert.DoesNotContain(".Visible", source);
        }

        private static string ReadAppSource(string appFileName)
        {
            string repoRoot = FindRepositoryRoot();
            return File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", appFileName));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
