using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRefreshCompletionDecisionServiceTests
    {
        [Fact]
        public void DecideCreatedCaseDisplayCompletion_AllowsOnlyDisplayCompletableRefreshAndForeground()
        {
            var service = new TaskPaneRefreshCompletionDecisionService();
            TaskPaneRefreshAttemptResult attempt = TaskPaneRefreshAttemptResult.Succeeded()
                .WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome.Completed(
                    "visible",
                    VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                    PaneVisibleSource.RefreshedShown,
                    workbookWindowEnsureStatus: null,
                    fullRecoveryAttempted: false,
                    fullRecoverySucceeded: null));

            CreatedCaseDisplayCompletionDecision decision = service.DecideCreatedCaseDisplayCompletion(
                new TaskPaneRefreshCompletionDecisionInput(
                    isCreatedCaseDisplayReason: true,
                    attempt));

            Assert.True(decision.CanComplete);
            Assert.Equal(string.Empty, decision.BlockedReason);
        }

        [Fact]
        public void DecideCreatedCaseDisplayCompletion_BlocksNonDisplayCompletableForeground()
        {
            var service = new TaskPaneRefreshCompletionDecisionService();
            TaskPaneRefreshAttemptResult attempt = TaskPaneRefreshAttemptResult.Succeeded()
                .WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome.Completed(
                    "visible",
                    VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                    PaneVisibleSource.RefreshedShown,
                    workbookWindowEnsureStatus: null,
                    fullRecoveryAttempted: false,
                    fullRecoverySucceeded: null))
                .WithForegroundGuaranteeOutcome(ForegroundGuaranteeOutcome.RequiredFailed(
                    ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                    "foregroundFailed"));

            CreatedCaseDisplayCompletionDecision decision = service.DecideCreatedCaseDisplayCompletion(
                new TaskPaneRefreshCompletionDecisionInput(
                    isCreatedCaseDisplayReason: true,
                    attempt));

            Assert.False(decision.CanComplete);
            Assert.Equal("foregroundGuaranteeDisplayCompletable=false", decision.BlockedReason);
        }

        [Fact]
        public void Source_DoesNotOwnSessionsEmitRetryTimerOrForegroundExecution()
        {
            string source = ReadAppSource("TaskPaneRefreshCompletionDecisionService.cs");

            Assert.Contains("!attemptResult.IsRefreshSucceeded", source);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable", source);
            Assert.Contains("!IsForegroundDisplayCompletableTerminalInput(attemptResult.ForegroundGuaranteeOutcome)", source);
            Assert.DoesNotContain("_createdCaseDisplaySessions", source);
            Assert.DoesNotContain("IsCompleted", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("case-display-completed", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
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
