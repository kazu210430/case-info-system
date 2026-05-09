using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRetryTimerLifecycle
    {
        private readonly List<WaitReadyRetryTimerRegistration> _waitReadyRetryTimers = new List<WaitReadyRetryTimerRegistration>();

        private System.Windows.Forms.Timer _pendingPaneRefreshTimer;
        private EventHandler _pendingPaneRefreshTickHandler;

        internal int WaitReadyRetryTimerCount
        {
            get
            {
                return _waitReadyRetryTimers.Count;
            }
        }

        internal bool HasPendingPaneRefreshTimer
        {
            get
            {
                return _pendingPaneRefreshTimer != null;
            }
        }

        internal bool IsPendingPaneRefreshTimerEnabled
        {
            get
            {
                return _pendingPaneRefreshTimer != null && _pendingPaneRefreshTimer.Enabled;
            }
        }

        internal int PendingPaneRefreshTimerRegistrationCount
        {
            get
            {
                return _pendingPaneRefreshTimer == null || _pendingPaneRefreshTickHandler == null ? 0 : 1;
            }
        }

        internal void ScheduleWaitReadyRetryTimer(int intervalMs, Action retryAction)
        {
            if (retryAction == null)
            {
                return;
            }

            System.Windows.Forms.Timer retryTimer = new System.Windows.Forms.Timer
            {
                Interval = intervalMs
            };

            WaitReadyRetryTimerRegistration registration = null;
            EventHandler tickHandler = (sender, args) =>
            {
                CancelWaitReadyRetryTimer(registration);
                retryAction();
            };

            registration = new WaitReadyRetryTimerRegistration(retryTimer, tickHandler);
            _waitReadyRetryTimers.Add(registration);
            retryTimer.Tick += tickHandler;
            retryTimer.Start();
        }

        internal void StopWaitReadyRetryTimers()
        {
            if (_waitReadyRetryTimers.Count == 0)
            {
                return;
            }

            foreach (WaitReadyRetryTimerRegistration registration in _waitReadyRetryTimers.ToArray())
            {
                CancelWaitReadyRetryTimer(registration);
            }
        }

        internal void StartPendingPaneRefreshTimer(int intervalMs, EventHandler tickHandler)
        {
            if (tickHandler == null)
            {
                throw new ArgumentNullException(nameof(tickHandler));
            }

            EnsurePendingPaneRefreshTimer();
            if (_pendingPaneRefreshTickHandler != tickHandler)
            {
                if (_pendingPaneRefreshTickHandler != null)
                {
                    _pendingPaneRefreshTimer.Tick -= _pendingPaneRefreshTickHandler;
                }

                _pendingPaneRefreshTickHandler = tickHandler;
                _pendingPaneRefreshTimer.Tick += _pendingPaneRefreshTickHandler;
            }

            _pendingPaneRefreshTimer.Interval = intervalMs;
            _pendingPaneRefreshTimer.Stop();
            _pendingPaneRefreshTimer.Start();
        }

        internal void StopPendingPaneRefreshTimer()
        {
            if (_pendingPaneRefreshTimer == null)
            {
                return;
            }

            _pendingPaneRefreshTimer.Stop();
            if (_pendingPaneRefreshTickHandler != null)
            {
                _pendingPaneRefreshTimer.Tick -= _pendingPaneRefreshTickHandler;
            }

            _pendingPaneRefreshTimer.Dispose();
            _pendingPaneRefreshTimer = null;
            _pendingPaneRefreshTickHandler = null;
        }

        internal void StopAllRetryTimers()
        {
            StopPendingPaneRefreshTimer();
            StopWaitReadyRetryTimers();
        }

        private void EnsurePendingPaneRefreshTimer()
        {
            if (_pendingPaneRefreshTimer != null)
            {
                return;
            }

            _pendingPaneRefreshTimer = new System.Windows.Forms.Timer();
        }

        private void CancelWaitReadyRetryTimer(WaitReadyRetryTimerRegistration registration)
        {
            if (registration == null)
            {
                return;
            }

            _waitReadyRetryTimers.Remove(registration);
            registration.Timer.Stop();
            registration.Timer.Tick -= registration.TickHandler;
            registration.Timer.Dispose();
        }

        private sealed class WaitReadyRetryTimerRegistration
        {
            internal WaitReadyRetryTimerRegistration(System.Windows.Forms.Timer timer, EventHandler tickHandler)
            {
                Timer = timer ?? throw new ArgumentNullException(nameof(timer));
                TickHandler = tickHandler ?? throw new ArgumentNullException(nameof(tickHandler));
            }

            internal System.Windows.Forms.Timer Timer { get; }

            internal EventHandler TickHandler { get; }
        }
    }
}
