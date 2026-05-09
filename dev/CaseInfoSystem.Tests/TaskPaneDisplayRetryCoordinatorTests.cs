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
}
