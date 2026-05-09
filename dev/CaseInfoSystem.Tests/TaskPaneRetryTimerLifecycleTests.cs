using System;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRetryTimerLifecycleTests
    {
        [Fact]
        public void ScheduleWaitReadyRetryTimer_WhenActionIsNull_DoesNotRegisterTimer()
        {
            var lifecycle = new TaskPaneRetryTimerLifecycle();

            lifecycle.ScheduleWaitReadyRetryTimer(80, retryAction: null);

            Assert.Equal(0, lifecycle.WaitReadyRetryTimerCount);
        }

        [Fact]
        public void StopWaitReadyRetryTimers_CancelsAllReadyShowRetryTimers()
        {
            var lifecycle = new TaskPaneRetryTimerLifecycle();

            lifecycle.ScheduleWaitReadyRetryTimer(80, () => { });
            lifecycle.ScheduleWaitReadyRetryTimer(80, () => { });

            Assert.Equal(2, lifecycle.WaitReadyRetryTimerCount);

            lifecycle.StopWaitReadyRetryTimers();

            Assert.Equal(0, lifecycle.WaitReadyRetryTimerCount);
        }

        [Fact]
        public void StartPendingPaneRefreshTimer_RestartsSingleOwnedTimerRegistration()
        {
            var lifecycle = new TaskPaneRetryTimerLifecycle();
            EventHandler firstTickHandler = (sender, args) => { };
            EventHandler secondTickHandler = (sender, args) => { };

            lifecycle.StartPendingPaneRefreshTimer(400, firstTickHandler);
            lifecycle.StartPendingPaneRefreshTimer(400, secondTickHandler);

            Assert.True(lifecycle.HasPendingPaneRefreshTimer);
            Assert.True(lifecycle.IsPendingPaneRefreshTimerEnabled);
            Assert.Equal(1, lifecycle.PendingPaneRefreshTimerRegistrationCount);

            lifecycle.StopPendingPaneRefreshTimer();

            Assert.False(lifecycle.HasPendingPaneRefreshTimer);
            Assert.False(lifecycle.IsPendingPaneRefreshTimerEnabled);
            Assert.Equal(0, lifecycle.PendingPaneRefreshTimerRegistrationCount);
        }

        [Fact]
        public void StopAllRetryTimers_CleansReadyShowAndPendingTimersTogether()
        {
            var lifecycle = new TaskPaneRetryTimerLifecycle();

            lifecycle.ScheduleWaitReadyRetryTimer(80, () => { });
            lifecycle.StartPendingPaneRefreshTimer(400, (sender, args) => { });

            lifecycle.StopAllRetryTimers();

            Assert.Equal(0, lifecycle.WaitReadyRetryTimerCount);
            Assert.False(lifecycle.HasPendingPaneRefreshTimer);
            Assert.False(lifecycle.IsPendingPaneRefreshTimerEnabled);
            Assert.Equal(0, lifecycle.PendingPaneRefreshTimerRegistrationCount);
        }
    }
}
