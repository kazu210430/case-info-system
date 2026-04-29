using System.Collections.Generic;
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
        public void ShowWhenReady_SchedulesFallbackAfterMaxAttempts()
        {
            var attempts = new List<int>();
            bool shown = false;
            bool fallbackScheduled = false;
            string fallbackReason = null;
            var coordinator = new TaskPaneDisplayRetryCoordinator(maxAttempts: 2);

            coordinator.ShowWhenReady(
                workbook: null,
                reason: "fallback",
                tryShowOnce: (_, __, attemptNumber) =>
                {
                    attempts.Add(attemptNumber);
                    return false;
                },
                scheduleRetry: (_, __, ___, continueAction) => continueAction(),
                onShown: () => shown = true,
                scheduleFallback: (_, reason) =>
                {
                    fallbackScheduled = true;
                    fallbackReason = reason;
                });

            Assert.Equal(new[] { 1, 2 }, attempts);
            Assert.False(shown);
            Assert.True(fallbackScheduled);
            Assert.Equal("fallback", fallbackReason);
        }
    }
}
