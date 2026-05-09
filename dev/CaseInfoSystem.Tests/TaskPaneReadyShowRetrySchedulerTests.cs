using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneReadyShowRetrySchedulerTests
    {
        [Fact]
        public void Schedule_WhenRetryActionIsNull_DoesNotRegisterTimer()
        {
            var logs = new List<string>();
            var lifecycle = new TaskPaneRetryTimerLifecycle();
            var scheduler = CreateScheduler(lifecycle, logs);

            scheduler.Schedule(null, "ready-show", 2, retryAction: null);

            Assert.Equal(0, lifecycle.WaitReadyRetryTimerCount);
            Assert.Contains(logs, log => log.Contains("action=wait-ready-retry-scheduled"));
            Assert.DoesNotContain(logs, log => log.Contains("action=wait-ready-retry-firing"));
        }

        [Fact]
        public void Schedule_RegistersReadyShowRetryTimerWithContractDelayAndAttempt()
        {
            var logs = new List<string>();
            var lifecycle = new TaskPaneRetryTimerLifecycle();
            var scheduler = CreateScheduler(lifecycle, logs);

            scheduler.Schedule(null, "ready-show", 2, () => { });

            Assert.Equal(1, lifecycle.WaitReadyRetryTimerCount);
            Assert.Contains(logs, log =>
                log.Contains("source=TaskPaneReadyShowRetryScheduler")
                && log.Contains("action=wait-ready-retry-scheduled")
                && log.Contains("attempt=2")
                && log.Contains("maxAttempts=2")
                && log.Contains("retryDelayMs=80"));

            lifecycle.StopWaitReadyRetryTimers();
        }

        private static TaskPaneReadyShowRetryScheduler CreateScheduler(
            TaskPaneRetryTimerLifecycle lifecycle,
            List<string> logs)
        {
            return new TaskPaneReadyShowRetryScheduler(
                new Logger(logs.Add),
                lifecycle,
                _ => "full=\"\",name=\"\"",
                _ => string.Empty);
        }
    }
}
