using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class PendingPaneRefreshRetryService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly ExcelInteropService _excelInteropService;
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly Logger _logger;
        private readonly int _pendingPaneRefreshIntervalMs;
        private readonly int _pendingPaneRefreshMaxAttempts;
        private readonly TaskPaneRetryTimerLifecycle _retryTimerLifecycle;
        private readonly Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> _tryRefreshTaskPane;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Action _stopPendingPaneRefreshTimer;
        private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
        private readonly Func<Excel.Window, string> _formatWindowDescriptor;
        private readonly Func<string> _formatActiveState;
        private readonly Func<Excel.Workbook, string> _safeWorkbookFullName;
        private readonly Func<Excel.Window, string> _safeWindowHwnd;
        private readonly PendingPaneRefreshRetryState _retryState = new PendingPaneRefreshRetryState();

        internal PendingPaneRefreshRetryService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            int pendingPaneRefreshIntervalMs,
            int pendingPaneRefreshMaxAttempts,
            TaskPaneRetryTimerLifecycle retryTimerLifecycle,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Action stopPendingPaneRefreshTimer,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<Excel.Window, string> formatWindowDescriptor,
            Func<string> formatActiveState,
            Func<Excel.Workbook, string> safeWorkbookFullName,
            Func<Excel.Window, string> safeWindowHwnd)
        {
            _excelInteropService = excelInteropService;
            _workbookSessionService = workbookSessionService;
            _logger = logger;
            _pendingPaneRefreshIntervalMs = pendingPaneRefreshIntervalMs;
            _pendingPaneRefreshMaxAttempts = pendingPaneRefreshMaxAttempts;
            _retryTimerLifecycle = retryTimerLifecycle ?? throw new ArgumentNullException(nameof(retryTimerLifecycle));
            _tryRefreshTaskPane = tryRefreshTaskPane ?? throw new ArgumentNullException(nameof(tryRefreshTaskPane));
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow ?? throw new ArgumentNullException(nameof(resolveWorkbookPaneWindow));
            _stopPendingPaneRefreshTimer = stopPendingPaneRefreshTimer ?? throw new ArgumentNullException(nameof(stopPendingPaneRefreshTimer));
            _formatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
            _formatWindowDescriptor = formatWindowDescriptor ?? throw new ArgumentNullException(nameof(formatWindowDescriptor));
            _formatActiveState = formatActiveState ?? throw new ArgumentNullException(nameof(formatActiveState));
            _safeWorkbookFullName = safeWorkbookFullName ?? throw new ArgumentNullException(nameof(safeWorkbookFullName));
            _safeWindowHwnd = safeWindowHwnd ?? throw new ArgumentNullException(nameof(safeWindowHwnd));
        }

        internal int AttemptsRemaining
        {
            get
            {
                return _retryState.AttemptsRemaining;
            }
        }

        internal void TrackActiveTarget()
        {
            _retryState.TrackActiveTarget();
        }

        internal void TrackWorkbookTarget(string workbookFullName)
        {
            _retryState.TrackWorkbookTarget(workbookFullName);
        }

        internal int BeginRetrySequence(string reason)
        {
            _retryState.BeginRetrySequence(reason, _pendingPaneRefreshMaxAttempts);
            _retryTimerLifecycle.StartPendingPaneRefreshTimer(
                _pendingPaneRefreshIntervalMs,
                PendingPaneRefreshTimer_Tick);
            return _retryState.AttemptsRemaining;
        }

        internal void StopTimer()
        {
            _retryTimerLifecycle.StopPendingPaneRefreshTimer();
        }

        private void PendingPaneRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (!_retryState.HasAttemptsRemaining)
            {
                ResolvePendingRetryContinuation(PendingRetryTickResult.StopRetrySequence());
                return;
            }

            PendingRetryTickResult tickResult = TryRefreshPendingWorkbookTarget();
            if (!tickResult.Handled)
            {
                tickResult = TryRefreshPendingActiveContextFallback();
            }

            ResolvePendingRetryContinuation(tickResult);
        }

        private PendingRetryTickResult TryRefreshPendingWorkbookTarget()
        {
            Excel.Workbook targetWorkbook = ResolvePendingPaneRefreshWorkbook();
            if (targetWorkbook == null)
            {
                return PendingRetryTickResult.ContinueToActiveContextFallback();
            }

            int attemptsRemaining = _retryState.ConsumeAttempt();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-retry-start reason="
                + _retryState.Reason
                + ", workbook="
                + _formatWorkbookDescriptor(targetWorkbook)
                + ", attemptsRemaining="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
            _logger?.Info("TaskPane timer retry start. reason=" + _retryState.Reason + ", workbook=" + _safeWorkbookFullName(targetWorkbook) + ", attemptsRemaining=" + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
            Excel.Window workbookWindow = _resolveWorkbookPaneWindow(targetWorkbook, _retryState.Reason, true);
            bool refreshed = _tryRefreshTaskPane(_retryState.Reason, targetWorkbook, workbookWindow).IsRefreshSucceeded;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-retry-end reason="
                + _retryState.Reason
                + ", workbook="
                + _formatWorkbookDescriptor(targetWorkbook)
                + ", resolvedWindow="
                + _formatWindowDescriptor(workbookWindow)
                + ", refreshed="
                + refreshed.ToString());
            _logger?.Info("TaskPane timer retry result. reason=" + _retryState.Reason + ", workbook=" + _safeWorkbookFullName(targetWorkbook) + ", windowHwnd=" + _safeWindowHwnd(workbookWindow) + ", refreshed=" + refreshed.ToString());
            return refreshed
                ? PendingRetryTickResult.StopRetrySequence()
                : PendingRetryTickResult.ContinueRetrySequence();
        }

        private PendingRetryTickResult TryRefreshPendingActiveContextFallback()
        {
            WorkbookContext context = _workbookSessionService == null ? null : _workbookSessionService.ResolveActiveContext("PendingPaneRefresh");
            if (context == null || context.Role != WorkbookRole.Case)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-stop reason="
                    + _retryState.Reason
                    + ", pendingWorkbookFullName=\""
                    + _retryState.WorkbookFullName
                    + "\", contextRole="
                    + (context == null ? "null" : context.Role.ToString())
                    + ", attemptsRemaining="
                    + _retryState.AttemptsRemaining.ToString(CultureInfo.InvariantCulture));
                return PendingRetryTickResult.StopRetrySequence();
            }

            int fallbackAttemptsRemaining = _retryState.ConsumeAttempt();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-start reason="
                + _retryState.Reason
                + ", pendingWorkbookFullName=\""
                + _retryState.WorkbookFullName
                + "\", contextWorkbook="
                + _formatWorkbookDescriptor(context.Workbook)
                + ", attemptsRemaining="
                + fallbackAttemptsRemaining.ToString(CultureInfo.InvariantCulture)
                + ", activeState="
                + _formatActiveState());
            _logger?.Info(
                "TaskPane timer fallback active CASE context start. reason="
                + _retryState.Reason
                + ", pendingWorkbook="
                + _retryState.WorkbookFullName
                + ", attemptsRemaining="
                + fallbackAttemptsRemaining.ToString(CultureInfo.InvariantCulture));
            bool fallbackRefreshed = _tryRefreshTaskPane(_retryState.Reason, null, null).IsRefreshSucceeded;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-end reason="
                + _retryState.Reason
                + ", pendingWorkbookFullName=\""
                + _retryState.WorkbookFullName
                + "\", contextWorkbook="
                + _formatWorkbookDescriptor(context.Workbook)
                + ", refreshed="
                + fallbackRefreshed.ToString()
                + ", activeState="
                + _formatActiveState());
            _logger?.Info(
                "TaskPane timer fallback active CASE context result. reason="
                + _retryState.Reason
                + ", pendingWorkbook="
                + _retryState.WorkbookFullName
                + ", refreshed="
                + fallbackRefreshed.ToString());
            return fallbackRefreshed
                ? PendingRetryTickResult.StopRetrySequence()
                : PendingRetryTickResult.ContinueRetrySequence();
        }

        private void ResolvePendingRetryContinuation(PendingRetryTickResult tickResult)
        {
            if (tickResult.ShouldStopTimer)
            {
                _stopPendingPaneRefreshTimer();
            }
        }

        private Excel.Workbook ResolvePendingPaneRefreshWorkbook()
        {
            if (_excelInteropService == null || string.IsNullOrWhiteSpace(_retryState.WorkbookFullName))
            {
                return null;
            }

            return _excelInteropService.FindOpenWorkbook(_retryState.WorkbookFullName);
        }

        private readonly struct PendingRetryTickResult
        {
            private PendingRetryTickResult(bool handled, bool shouldStopTimer)
            {
                Handled = handled;
                ShouldStopTimer = shouldStopTimer;
            }

            internal bool Handled { get; }

            internal bool ShouldStopTimer { get; }

            internal static PendingRetryTickResult ContinueToActiveContextFallback()
            {
                return new PendingRetryTickResult(handled: false, shouldStopTimer: false);
            }

            internal static PendingRetryTickResult ContinueRetrySequence()
            {
                return new PendingRetryTickResult(handled: true, shouldStopTimer: false);
            }

            internal static PendingRetryTickResult StopRetrySequence()
            {
                return new PendingRetryTickResult(handled: true, shouldStopTimer: true);
            }
        }

        private sealed class PendingPaneRefreshRetryState
        {
            internal int AttemptsRemaining { get; private set; }

            internal string Reason { get; private set; } = string.Empty;

            internal string WorkbookFullName { get; private set; } = string.Empty;

            internal bool HasAttemptsRemaining
            {
                get
                {
                    return AttemptsRemaining > 0;
                }
            }

            internal void TrackActiveTarget()
            {
                WorkbookFullName = string.Empty;
            }

            internal void TrackWorkbookTarget(string workbookFullName)
            {
                WorkbookFullName = workbookFullName ?? string.Empty;
            }

            internal void BeginRetrySequence(string reason, int maxAttempts)
            {
                Reason = reason ?? string.Empty;
                AttemptsRemaining = maxAttempts;
            }

            internal int ConsumeAttempt()
            {
                if (AttemptsRemaining <= 0)
                {
                    return 0;
                }

                AttemptsRemaining--;
                return AttemptsRemaining;
            }
        }
    }
}
