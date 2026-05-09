using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneDisplayRetryCoordinatorTests
    {
        [Fact]
        public void ShowWhenReady_RetriesThroughSchedulerUntilAttemptSucceeds()
        {
            var attempts = new List<int>();
            var scheduledAttempts = new List<int>();
            bool shown = false;
            bool fallbackScheduled = false;
            var coordinator = new TaskPaneDisplayRetryCoordinator(maxAttempts: 3);

            coordinator.ShowWhenReady(
                workbook: null,
                reason: "test",
                tryShowOnce: (_, __, attemptNumber) =>
                {
                    attempts.Add(attemptNumber);
                    return attemptNumber == 3;
                },
                scheduleRetry: (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                onShown: () => shown = true,
                scheduleFallback: (_, __) => fallbackScheduled = true);

            Assert.Equal(new[] { 1, 2, 3 }, attempts);
            Assert.Equal(new[] { 2, 3 }, scheduledAttempts);
            Assert.True(shown);
            Assert.False(fallbackScheduled);
        }

        [Fact]
        public void ShowWhenReady_WhenReadyShowAttemptsExhausted_SchedulesFallbackWithoutAttempt3()
        {
            var attempts = new List<int>();
            var scheduledAttempts = new List<int>();
            bool shown = false;
            bool fallbackScheduled = false;
            string fallbackReason = null;
            var coordinator = new TaskPaneDisplayRetryCoordinator(WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts);

            coordinator.ShowWhenReady(
                workbook: null,
                reason: "fallback",
                tryShowOnce: (_, __, attemptNumber) =>
                {
                    attempts.Add(attemptNumber);
                    return false;
                },
                scheduleRetry: (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                onShown: () => shown = true,
                scheduleFallback: (_, reason) =>
                {
                    fallbackScheduled = true;
                    fallbackReason = reason;
                });

            Assert.Equal(new[] { 1, 2 }, attempts);
            Assert.Equal(new[] { 2 }, scheduledAttempts);
            Assert.DoesNotContain(3, attempts);
            Assert.DoesNotContain(3, scheduledAttempts);
            Assert.False(shown);
            Assert.True(fallbackScheduled);
            Assert.Equal("fallback", fallbackReason);
        }

        [Fact]
        public void ReadyShowRetryContract_UsesTwoAttemptsSeparateFromPendingRetry()
        {
            Assert.Equal(2, WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts);
            Assert.Equal(80, WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowRetryDelayMs);

            string repoRoot = FindRepositoryRoot();
            string compositionSource = File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "AddInCompositionRoot.cs"));
            string orchestrationSource = File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", "TaskPaneRefreshOrchestrationService.cs"));
            string thisAddInSource = File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "ThisAddIn.cs"));

            Assert.Contains("new TaskPaneDisplayRetryCoordinator(ReadyShowRetryMaxAttempts)", compositionSource);
            Assert.DoesNotContain("new TaskPaneDisplayRetryCoordinator(_pendingPaneRefreshMaxAttempts)", compositionSource);
            Assert.Contains("new TaskPaneReadyShowRetryScheduler", orchestrationSource);
            Assert.DoesNotContain("private void ScheduleTaskPaneReadyRetry", orchestrationSource);
            Assert.DoesNotContain("TaskPaneRefreshOrchestrationService.PendingPaneRefreshMaxAttempts", thisAddInSource);
            Assert.Contains("internal const int PendingPaneRefreshIntervalMs = 400;", orchestrationSource);
            Assert.Contains("internal const int PendingPaneRefreshMaxAttempts = 3;", orchestrationSource);
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

    public sealed class DisplayConvergenceCompletionBoundarySourceTests
    {
        [Fact]
        public void CaseDisplayCompleted_EmitOwnerAndOneTimeGateStayInOrchestrationOnly()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");

            Assert.Contains("TryCompleteCreatedCaseDisplaySession", orchestrationSource);
            Assert.Contains("action=case-display-completed", orchestrationSource);
            Assert.Contains("\"case-display-completed\"", orchestrationSource);
            Assert.Contains("\"TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession\"", orchestrationSource);
            Assert.Contains("NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);", orchestrationSource);
            AssertContainsInOrder(
                orchestrationSource,
                "bool shouldEmit = false;",
                "if (!resolvedSession.IsCompleted)",
                "resolvedSession.IsCompleted = true;",
                "_createdCaseDisplaySessions.Remove(resolvedSession.WorkbookFullName);",
                "shouldEmit = true;",
                "if (!shouldEmit)",
                "return;",
                "string details =");

            AssertDoesNotOwnCompletion("WorkbookTaskPaneReadyShowAttemptWorker.cs");
            AssertDoesNotOwnCompletion("PendingPaneRefreshRetryService.cs");
            AssertDoesNotOwnCompletion("TaskPaneDisplayRetryCoordinator.cs");
            AssertDoesNotOwnCompletion("TaskPaneRefreshCoordinator.cs");
            AssertDoesNotOwnCompletion("WindowActivatePaneHandlingService.cs");
            AssertDoesNotOwnCompletion("WindowActivateDownstreamObservation.cs");
        }

        [Fact]
        public void CompletionHardGateDecisionContract_RequiresVisibilityAndForegroundDisplayCompletableFacts()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");

            Assert.Contains("IsCreatedCaseDisplayReason(reason)", orchestrationSource);
            Assert.Contains("attemptResult == null", orchestrationSource);
            Assert.Contains("!attemptResult.IsRefreshSucceeded", orchestrationSource);
            Assert.Contains("!attemptResult.IsPaneVisible", orchestrationSource);
            Assert.Contains("attemptResult.VisibilityRecoveryOutcome == null", orchestrationSource);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsTerminal", orchestrationSource);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable", orchestrationSource);
            Assert.Contains("!attemptResult.IsForegroundGuaranteeTerminal", orchestrationSource);
            Assert.Contains("attemptResult.ForegroundGuaranteeOutcome == null", orchestrationSource);
            Assert.Contains("!attemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable", orchestrationSource);
        }

        [Fact]
        public void ReadyShowCallback_ReturnsFactsToConvergenceChainBeforeCompletionGate()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string readyShowHandoff = Slice(
                orchestrationSource,
                "CreatedCaseDisplaySession createdCaseDisplaySession = BeginCreatedCaseDisplaySession",
                "internal Excel.Window ResolveWorkbookPaneWindow");
            string callbackHandler = Slice(
                orchestrationSource,
                "private void HandleWorkbookTaskPaneShown",
                "private void TryCompleteCreatedCaseDisplaySession");
            string workerSource = ReadAppSource("WorkbookTaskPaneReadyShowAttemptWorker.cs");

            Assert.Contains(
                "outcome => HandleWorkbookTaskPaneShown(createdCaseDisplaySession, workbook, reason, outcome)",
                readyShowHandoff);
            AssertContainsInOrder(
                callbackHandler,
                "CompleteVisibilityRecoveryOutcome(",
                "CompleteRefreshSourceSelectionOutcome(",
                "CompleteRebuildFallbackOutcome(",
                "CompleteForegroundGuaranteeOutcome(",
                "TryCompleteCreatedCaseDisplaySession(");
            Assert.Contains("() => onShown?.Invoke(shownOutcome)", workerSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", workerSource);
            Assert.DoesNotContain("case-display-completed", workerSource);
        }

        [Fact]
        public void PendingRetryAndActiveFallbackRefreshSuccessStopRetryWithoutCompletionOwnership()
        {
            string pendingSource = ReadAppSource("PendingPaneRefreshRetryService.cs");

            AssertContainsInOrder(
                pendingSource,
                "action=defer-retry-end",
                "refreshed=",
                "if (refreshed)",
                "_stopPendingPaneRefreshTimer();");
            AssertContainsInOrder(
                pendingSource,
                "action=defer-active-context-fallback-end",
                "refreshed=",
                "if (fallbackRefreshed)",
                "_stopPendingPaneRefreshTimer();");
            Assert.DoesNotContain("case-display-completed", pendingSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", pendingSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", pendingSource);
        }

        [Fact]
        public void ForegroundAndNormalizedOutcomesDoNotOwnCompletionEmit()
        {
            string attemptResultSource = ReadAppSource("TaskPaneRefreshAttemptResult.cs");
            string normalizedMapperSource = ReadAppSource("TaskPaneNormalizedOutcomeMapper.cs");
            string requiredDegraded = Slice(
                attemptResultSource,
                "internal static ForegroundGuaranteeOutcome RequiredDegraded",
                "internal static ForegroundGuaranteeOutcome RequiredFailed");

            Assert.Contains("ForegroundGuaranteeOutcomeStatus.RequiredDegraded", requiredDegraded);
            Assert.Contains("isTerminal: true", requiredDegraded);
            Assert.Contains("isDisplayCompletable: true", requiredDegraded);
            Assert.Contains("recoverySucceeded: false", requiredDegraded);
            Assert.DoesNotContain("case-display-completed", attemptResultSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", attemptResultSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", attemptResultSource);
            Assert.DoesNotContain("case-display-completed", normalizedMapperSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", normalizedMapperSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", normalizedMapperSource);
        }

        private static void AssertDoesNotOwnCompletion(string appFileName)
        {
            string source = ReadAppSource(appFileName);
            Assert.DoesNotContain("action=case-display-completed", source);
            Assert.DoesNotContain("\"case-display-completed\"", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", source);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", source);
        }

        private static void AssertContainsInOrder(string source, params string[] fragments)
        {
            int previousIndex = -1;
            foreach (string fragment in fragments)
            {
                int index = source.IndexOf(fragment, previousIndex + 1, StringComparison.Ordinal);
                Assert.True(
                    index > previousIndex,
                    "Expected to find '" + fragment + "' after index " + previousIndex.ToString() + ".");
                previousIndex = index;
            }
        }

        private static string Slice(string source, string startFragment, string endFragment)
        {
            int start = source.IndexOf(startFragment, StringComparison.Ordinal);
            Assert.True(start >= 0, "Expected start fragment was not found: " + startFragment);
            int end = source.IndexOf(endFragment, start + startFragment.Length, StringComparison.Ordinal);
            Assert.True(end > start, "Expected end fragment was not found: " + endFragment);
            return source.Substring(start, end - start);
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
