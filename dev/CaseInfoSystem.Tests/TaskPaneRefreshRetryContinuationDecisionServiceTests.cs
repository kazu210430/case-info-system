using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRefreshRetryContinuationDecisionServiceTests
    {
        [Fact]
        public void DecideBeforeTick_WhenAttemptsAreExhausted_StopsRetry()
        {
            var service = new TaskPaneRefreshRetryContinuationDecisionService();

            TaskPaneRefreshRetryContinuationDecision decision = service.DecideBeforeTick(hasAttemptsRemaining: false);

            Assert.True(decision.Handled);
            Assert.True(decision.ShouldStopTimer);
            Assert.False(decision.ShouldAttemptActiveContextFallback);
            Assert.Equal("stop-retry", decision.Action);
            Assert.Equal("attemptsExhausted", decision.Description);
        }

        [Fact]
        public void DecideAfterWorkbookTargetResolution_WhenTargetMissing_DefersToActiveContextFallback()
        {
            var service = new TaskPaneRefreshRetryContinuationDecisionService();

            TaskPaneRefreshRetryContinuationDecision decision =
                service.DecideAfterWorkbookTargetResolution(hasTargetWorkbook: false);

            Assert.False(decision.Handled);
            Assert.False(decision.ShouldStopTimer);
            Assert.Equal("continue-to-active-context-fallback", decision.Action);
            Assert.Equal("workbookTargetMissing", decision.Description);
        }

        [Fact]
        public void DecideActiveContextFallback_WhenActiveContextIsCase_AttemptsFallback()
        {
            var service = new TaskPaneRefreshRetryContinuationDecisionService();
            var context = new WorkbookContext(
                new Excel.Workbook(),
                new Excel.Window(),
                WorkbookRole.Case,
                "root",
                "path",
                "sheet");

            TaskPaneRefreshRetryContinuationDecision decision = service.DecideActiveContextFallback(context);

            Assert.True(decision.Handled);
            Assert.False(decision.ShouldStopTimer);
            Assert.True(decision.ShouldAttemptActiveContextFallback);
            Assert.Equal("attempt-active-context-fallback", decision.Action);
            Assert.Equal("activeCaseContext", decision.Description);
        }

        [Fact]
        public void DecideAfterRefresh_StopsOnlyWhenRefreshSucceeded()
        {
            var service = new TaskPaneRefreshRetryContinuationDecisionService();

            TaskPaneRefreshRetryContinuationDecision success = service.DecideAfterRefresh(refreshed: true);
            TaskPaneRefreshRetryContinuationDecision failure = service.DecideAfterRefresh(refreshed: false);

            Assert.True(success.ShouldStopTimer);
            Assert.Equal("refreshSucceeded", success.Description);
            Assert.False(failure.ShouldStopTimer);
            Assert.Equal("continue-retry", failure.Action);
            Assert.Equal("refreshFailed", failure.Description);
        }

        [Fact]
        public void Source_DoesNotOwnTimerCallbackCompletionOrDisplayExecution()
        {
            string source = ReadAppSource("TaskPaneRefreshRetryContinuationDecisionService.cs");

            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("StartPendingPaneRefreshTimer", source);
            Assert.DoesNotContain("StopPendingPaneRefreshTimer", source);
            Assert.DoesNotContain("TaskPaneReadyShowRetryScheduler", source);
            Assert.DoesNotContain("WindowActivateDownstreamObservation", source);
            Assert.DoesNotContain("TryRefreshTaskPane", source);
            Assert.DoesNotContain("ResolveWorkbookPaneWindow", source);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", source);
            Assert.DoesNotContain("case-display-completed", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("TaskPaneRefreshCoordinator", source);
            Assert.DoesNotContain("TaskPaneManager", source);
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
