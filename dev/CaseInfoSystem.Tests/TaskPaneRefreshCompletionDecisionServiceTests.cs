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
                BuildContext(isCreatedCaseDisplayReason: true, attempt));

            Assert.True(decision.CanComplete);
            Assert.True(decision.ShouldResolveSession);
            Assert.Equal(TaskPaneRefreshCompletionDecisionStatus.ReadyForSession, decision.Status);
            Assert.Equal(string.Empty, decision.BlockedReason);
            Assert.NotNull(decision.Material);
            Assert.True(decision.Material.IsForegroundDisplayCompletableTerminalInput);
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
                BuildContext(isCreatedCaseDisplayReason: true, attempt));

            Assert.False(decision.CanComplete);
            Assert.False(decision.ShouldResolveSession);
            Assert.Equal(TaskPaneRefreshCompletionDecisionStatus.Blocked, decision.Status);
            Assert.Equal("foregroundGuaranteeDisplayCompletable=false", decision.BlockedReason);
        }

        [Fact]
        public void ClassifyCreatedCaseDisplayCompletionResult_PreservesStatusReasonAndEmitMaterial()
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
            TaskPaneRefreshCompletionContext context = BuildContext(isCreatedCaseDisplayReason: true, attempt);
            CreatedCaseDisplayCompletionDecision decision = service.DecideCreatedCaseDisplayCompletion(context);
            var session = new CreatedCaseDisplaySessionSnapshot(
                "CDS-0001",
                @"C:\Cases\A.xlsx",
                "created-case-post-release",
                isCompleted: false);

            TaskPaneRefreshCompletionResult result =
                service.ClassifyCreatedCaseDisplayCompletionResult(context, decision, session);

            Assert.True(result.CanEmit);
            Assert.Equal(TaskPaneRefreshCompletionResultStatus.ReadyToEmit, result.Status);
            Assert.Equal("readyToEmit", result.ResultReason);
            Assert.Equal("created-case-post-release", result.Reason);
            Assert.Equal("ready-show-attempt", result.CompletionSource);
            Assert.Equal(2, result.AttemptNumber);
            Assert.Same(attempt, result.AttemptResult);
            Assert.Same(session, result.SessionSnapshot);
            Assert.Equal("CDS-0001", result.SessionId);
            Assert.Equal(@"C:\Cases\A.xlsx", result.WorkbookFullName);
        }

        [Fact]
        public void ClassifyCreatedCaseDisplayCompletionResult_BlocksMissingSessionWithoutEmit()
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
            TaskPaneRefreshCompletionContext context = BuildContext(isCreatedCaseDisplayReason: true, attempt);
            CreatedCaseDisplayCompletionDecision decision = service.DecideCreatedCaseDisplayCompletion(context);

            TaskPaneRefreshCompletionResult result =
                service.ClassifyCreatedCaseDisplayCompletionResult(context, decision, sessionSnapshot: null);

            Assert.False(result.CanEmit);
            Assert.Equal(TaskPaneRefreshCompletionResultStatus.SessionMissing, result.Status);
            Assert.Equal("session=null", result.ResultReason);
            Assert.Equal(string.Empty, result.SessionId);
            Assert.Equal(string.Empty, result.WorkbookFullName);
        }

        [Fact]
        public void Source_DoesNotOwnSessionsEmitRetryTimerOrForegroundExecution()
        {
            string source = ReadAppSource("TaskPaneRefreshCompletionDecisionService.cs");

            Assert.Contains("TaskPaneRefreshCompletionContext", source);
            Assert.Contains("TaskPaneRefreshCompletionMaterial.FromContext", source);
            Assert.Contains("TaskPaneRefreshCompletionResult", source);
            Assert.Contains("!material.IsRefreshSucceeded", source);
            Assert.Contains("!material.IsVisibilityRecoveryDisplayCompletable", source);
            Assert.Contains("!material.IsForegroundDisplayCompletableTerminalInput", source);
            Assert.DoesNotContain("_createdCaseDisplaySessions", source);
            Assert.DoesNotContain("IsCompleted", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("case-display-completed", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
        }

        private static TaskPaneRefreshCompletionContext BuildContext(
            bool isCreatedCaseDisplayReason,
            TaskPaneRefreshAttemptResult attempt)
        {
            return TaskPaneRefreshCompletionContext.FromInput(
                new TaskPaneRefreshCompletionContextInput(
                    "created-case-post-release",
                    isCreatedCaseDisplayReason: isCreatedCaseDisplayReason,
                    attemptResult: attempt,
                    completionSource: "ready-show-attempt",
                    attemptNumber: 2,
                    displayRequest: null,
                    workbook: null,
                    window: null));
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
