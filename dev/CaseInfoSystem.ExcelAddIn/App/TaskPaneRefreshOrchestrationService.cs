using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshOrchestrationService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        internal const int PendingPaneRefreshIntervalMs = 400;
        internal const int PendingPaneRefreshMaxAttempts = 3;
        internal const int WorkbookPaneWindowResolveAttempts = 2;
        internal const int WorkbookPaneWindowResolveDelayMs = 80;

        private readonly ExcelInteropService _excelInteropService;
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly Logger _logger;
        private readonly TaskPaneRefreshCoordinator _taskPaneRefreshCoordinator;
        private readonly WorkbookTaskPaneReadyShowAttemptWorker _workbookTaskPaneReadyShowAttemptWorker;
        private readonly WorkbookPaneWindowResolver _workbookPaneWindowResolver;
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly ICasePaneHostBridge _casePaneHostBridge;
        private readonly PendingPaneRefreshRetryService _pendingPaneRefreshRetryService;

        private readonly List<System.Windows.Forms.Timer> _waitReadyRetryTimers = new List<System.Windows.Forms.Timer>();
        private int _kernelFlickerTraceRefreshAttemptSequence;

        internal TaskPaneRefreshOrchestrationService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            TaskPaneRefreshCoordinator taskPaneRefreshCoordinator,
            WorkbookTaskPaneReadyShowAttemptWorker workbookTaskPaneReadyShowAttemptWorker,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount,
            ICasePaneHostBridge casePaneHostBridge)
        {
            _excelInteropService = excelInteropService;
            _workbookSessionService = workbookSessionService;
            _logger = logger;
            _taskPaneRefreshCoordinator = taskPaneRefreshCoordinator;
            _workbookTaskPaneReadyShowAttemptWorker = workbookTaskPaneReadyShowAttemptWorker ?? throw new ArgumentNullException(nameof(workbookTaskPaneReadyShowAttemptWorker));
            _workbookPaneWindowResolver = new WorkbookPaneWindowResolver(
                _excelInteropService,
                _logger,
                workbook => FormatWorkbookDescriptor(workbook),
                window => FormatWindowDescriptor(window),
                () => FormatActiveState());
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
            _casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException(nameof(casePaneHostBridge));
            _pendingPaneRefreshRetryService = new PendingPaneRefreshRetryService(
                _excelInteropService,
                _workbookSessionService,
                _logger,
                PendingPaneRefreshIntervalMs,
                PendingPaneRefreshMaxAttempts,
                TryRefreshTaskPane,
                ResolveWorkbookPaneWindow,
                StopPendingPaneRefreshTimer,
                workbook => FormatWorkbookDescriptor(workbook),
                window => FormatWindowDescriptor(window),
                () => FormatActiveState(),
                workbook => SafeWorkbookFullName(workbook),
                window => SafeWindowHwnd(window));
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            int refreshAttemptId = ++_kernelFlickerTraceRefreshAttemptSequence;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=try-refresh-start refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", activeState="
                + FormatActiveState());
            RefreshPreconditionEvaluationResult preconditionEvaluationResult = RefreshPreconditionEvaluator.Evaluate(reason, workbook, window, _casePaneHostBridge);
            if (!preconditionEvaluationResult.CanRefresh)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action="
                    + preconditionEvaluationResult.SkipActionName
                    + " refreshAttemptId="
                    + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", inputWindow="
                    + FormatWindowDescriptor(window)
                    + ", activeState="
                    + FormatActiveState());
                return TaskPaneRefreshAttemptResult.Skipped();
            }

            RefreshDispatchExecutionResult dispatchExecutionResult = RefreshDispatchShell.Dispatch(
                _taskPaneRefreshCoordinator,
                reason,
                workbook,
                window,
                _getKernelHomeForm,
                _getTaskPaneRefreshSuppressionCount);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=try-refresh-end refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", result="
                + dispatchExecutionResult.ResultText);
            return dispatchExecutionResult.AttemptResult;
        }

        internal bool IsTaskPaneRefreshSucceeded(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return TryRefreshTaskPane(reason, workbook, window).IsRefreshSucceeded;
        }

        internal void RefreshActiveTaskPane(string reason)
        {
            TryRefreshTaskPane(reason, null, null);
        }

        internal void ScheduleActiveTaskPaneRefresh(string reason)
        {
            _pendingPaneRefreshRetryService.TrackActiveTarget();
            if (IsTaskPaneRefreshSucceeded(reason, null, null))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-immediate-success reason="
                    + (reason ?? string.Empty)
                    + ", target=active");
                StopPendingPaneRefreshTimer();
                return;
            }

            int attemptsRemaining = _pendingPaneRefreshRetryService.BeginRetrySequence(reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (reason ?? string.Empty)
                + ", target=active"
                + ", attempts="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=wait-ready-fallback-handoff reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", maxAttempts="
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true"
                + ", activeState="
                + FormatActiveState());
            _logger?.Info(
                "TaskPane wait-ready fallback handoff. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", maxAttempts="
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true");
            if (TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window: null))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=skip-workbook-open-defer reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveState());
                return;
            }

            _pendingPaneRefreshRetryService.TrackWorkbookTarget(_excelInteropService == null
                ? string.Empty
                : _excelInteropService.GetWorkbookFullName(workbook));
            Excel.Window workbookWindow = ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook: false);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-prepare reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", windowResolved="
                + (workbookWindow != null).ToString()
                + ", resolvedWindow="
                + FormatWindowDescriptor(workbookWindow)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true"
                + ", activeState="
                + FormatActiveState());
            _logger?.Info(
                "TaskPane timer fallback prepare. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", windowResolved="
                + (workbookWindow != null).ToString()
                + ", resolvedWindowHwnd="
                + SafeWindowHwnd(workbookWindow)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true");

            if (TryRefreshTaskPane(reason, workbook, workbookWindow).IsRefreshSucceeded)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-immediate-success reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook));
                _logger?.Info("TaskPane timer fallback immediate refresh succeeded. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook));
                StopPendingPaneRefreshTimer();
                return;
            }

            int attemptsRemaining = _pendingPaneRefreshRetryService.BeginRetrySequence(reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", attempts="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
            _logger?.Info("TaskPane timer fallback scheduled. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempts=" + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        internal void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            _workbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(
                workbook,
                reason,
                ScheduleTaskPaneReadyRetry,
                StopPendingPaneRefreshTimer,
                ScheduleWorkbookTaskPaneRefresh);
        }

        internal Excel.Window ResolveWorkbookPaneWindow(Excel.Workbook workbook, string reason, bool activateWorkbook)
        {
            return _workbookPaneWindowResolver.Resolve(workbook, reason, activateWorkbook);
        }

        internal void StopPendingPaneRefreshTimer()
        {
            _pendingPaneRefreshRetryService.StopTimer();
            StopWaitReadyRetryTimers();
        }

        private void ScheduleTaskPaneReadyRetry(Excel.Workbook workbook, string reason, int attemptNumber, Action retryAction)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=wait-ready-retry-scheduled reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryScheduled=true"
                + ", retryDelayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture)
                + ", delayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
            _logger?.Info(
                "TaskPane wait-ready retry scheduled. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryScheduled=true"
                + ", retryDelayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));

            if (retryAction == null)
            {
                return;
            }

            System.Windows.Forms.Timer retryTimer = new System.Windows.Forms.Timer
            {
                Interval = WorkbookPaneWindowResolveDelayMs
            };

            EventHandler tickHandler = null;
            tickHandler = (sender, args) =>
            {
                retryTimer.Stop();
                retryTimer.Tick -= tickHandler;
                _waitReadyRetryTimers.Remove(retryTimer);
                retryTimer.Dispose();
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=wait-ready-retry-firing reason="
                    + (reason ?? string.Empty)
                    + ", readyShowReason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", attempt="
                    + attemptNumber.ToString(CultureInfo.InvariantCulture)
                    + ", maxAttempts="
                    + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                    + ", retryDelayMs="
                    + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
                _logger?.Info(
                    "TaskPane wait-ready retry firing. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + SafeWorkbookFullName(workbook)
                    + ", readyShowReason="
                    + (reason ?? string.Empty)
                    + ", attempt="
                    + attemptNumber.ToString(CultureInfo.InvariantCulture)
                    + ", maxAttempts="
                    + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                    + ", retryDelayMs="
                    + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
                retryAction();
            };

            _waitReadyRetryTimers.Add(retryTimer);
            retryTimer.Tick += tickHandler;
            retryTimer.Start();
        }

        private void StopWaitReadyRetryTimers()
        {
            if (_waitReadyRetryTimers.Count == 0)
            {
                return;
            }

            foreach (System.Windows.Forms.Timer retryTimer in _waitReadyRetryTimers.ToArray())
            {
                retryTimer.Stop();
                retryTimer.Dispose();
            }

            _waitReadyRetryTimers.Clear();
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        private static class RefreshPreconditionEvaluator
        {
            internal static RefreshPreconditionEvaluationResult Evaluate(string reason, Excel.Workbook workbook, Excel.Window window, ICasePaneHostBridge casePaneHostBridge)
            {
                if (TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window))
                {
                    return RefreshPreconditionEvaluationResult.SkipWorkbookOpenWindowDependentRefresh();
                }

                if (casePaneHostBridge.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window))
                {
                    return RefreshPreconditionEvaluationResult.IgnoreDuringProtection();
                }

                return RefreshPreconditionEvaluationResult.Proceed();
            }
        }

        private static class RefreshDispatchShell
        {
            internal static RefreshDispatchExecutionResult Dispatch(
                TaskPaneRefreshCoordinator taskPaneRefreshCoordinator,
                string reason,
                Excel.Workbook workbook,
                Excel.Window window,
                Func<KernelHomeForm> getKernelHomeForm,
                Func<int> getTaskPaneRefreshSuppressionCount)
            {
                TaskPaneRefreshAttemptResult attemptResult = taskPaneRefreshCoordinator.TryRefreshTaskPane(
                    reason,
                    workbook,
                    window,
                    getKernelHomeForm == null ? null : getKernelHomeForm(),
                    getTaskPaneRefreshSuppressionCount == null ? 0 : getTaskPaneRefreshSuppressionCount());
                return RefreshDispatchExecutionResult.FromAttemptResult(attemptResult);
            }
        }

        private sealed class RefreshDispatchExecutionResult
        {
            private RefreshDispatchExecutionResult(TaskPaneRefreshAttemptResult attemptResult, string resultText)
            {
                AttemptResult = attemptResult;
                ResultText = resultText ?? string.Empty;
            }

            internal TaskPaneRefreshAttemptResult AttemptResult { get; }

            internal string ResultText { get; }

            internal static RefreshDispatchExecutionResult FromAttemptResult(TaskPaneRefreshAttemptResult attemptResult)
            {
                return new RefreshDispatchExecutionResult(
                    attemptResult,
                    attemptResult == null ? "null" : attemptResult.IsRefreshSucceeded.ToString());
            }
        }

        private sealed class RefreshPreconditionEvaluationResult
        {
            private RefreshPreconditionEvaluationResult(bool canRefresh, string skipActionName)
            {
                CanRefresh = canRefresh;
                SkipActionName = skipActionName ?? string.Empty;
            }

            internal bool CanRefresh { get; }

            internal string SkipActionName { get; }

            internal static RefreshPreconditionEvaluationResult Proceed()
            {
                return new RefreshPreconditionEvaluationResult(true, string.Empty);
            }

            internal static RefreshPreconditionEvaluationResult SkipWorkbookOpenWindowDependentRefresh()
            {
                return new RefreshPreconditionEvaluationResult(false, "skip-workbook-open-window-dependent-refresh");
            }

            internal static RefreshPreconditionEvaluationResult IgnoreDuringProtection()
            {
                return new RefreshPreconditionEvaluationResult(false, "ignore-during-protection");
            }
        }

        private sealed class WorkbookPaneWindowResolver
        {
            private readonly ExcelInteropService _excelInteropService;
            private readonly Logger _logger;
            private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
            private readonly Func<Excel.Window, string> _formatWindowDescriptor;
            private readonly Func<string> _formatActiveState;

            internal WorkbookPaneWindowResolver(
                ExcelInteropService excelInteropService,
                Logger logger,
                Func<Excel.Workbook, string> formatWorkbookDescriptor,
                Func<Excel.Window, string> formatWindowDescriptor,
                Func<string> formatActiveState)
            {
                _excelInteropService = excelInteropService;
                _logger = logger;
                _formatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
                _formatWindowDescriptor = formatWindowDescriptor ?? throw new ArgumentNullException(nameof(formatWindowDescriptor));
                _formatActiveState = formatActiveState ?? throw new ArgumentNullException(nameof(formatActiveState));
            }

            internal Excel.Window Resolve(Excel.Workbook workbook, string reason, bool activateWorkbook)
            {
                if (_excelInteropService == null || workbook == null)
                {
                    return null;
                }

                string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=resolve-window-start reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook)
                    + ", activateWorkbook="
                    + activateWorkbook.ToString()
                    + ", activeState="
                    + _formatActiveState());
                for (int attempt = 0; attempt < WorkbookPaneWindowResolveAttempts; attempt++)
                {
                    if (activateWorkbook)
                    {
                        _excelInteropService.ActivateWorkbook(workbook);
                    }

                    Excel.Window workbookWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
                    Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();
                    string activeWorkbookFullName = _excelInteropService.GetWorkbookFullName(activeWorkbook);
                    Excel.Window activeWindow = _excelInteropService.GetActiveWindow();
                    bool activeWorkbookMatches = string.Equals(activeWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase);

                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneRefreshOrchestrationService action=resolve-window-state reason="
                        + (reason ?? string.Empty)
                        + ", workbookFullName=\""
                        + workbookFullName
                        + "\", resolveAttempt="
                        + (attempt + 1).ToString(CultureInfo.InvariantCulture)
                        + ", activateWorkbook="
                        + activateWorkbook.ToString()
                        + ", visibleWindow="
                        + _formatWindowDescriptor(workbookWindow)
                        + ", activeWorkbook="
                        + _formatWorkbookDescriptor(activeWorkbook)
                        + ", activeWindow="
                        + _formatWindowDescriptor(activeWindow)
                        + ", activeWorkbookMatches="
                        + activeWorkbookMatches.ToString());
                    _logger?.Info("ResolveWorkbookPaneWindow state. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", resolveAttempt=" + (attempt + 1).ToString(CultureInfo.InvariantCulture) + ", activateWorkbook=" + activateWorkbook.ToString() + ", visibleWindowHwnd=" + SafeWindowHwnd(workbookWindow) + ", activeWorkbook=" + activeWorkbookFullName + ", activeWorkbookMatches=" + activeWorkbookMatches.ToString() + ", activeWindowHwnd=" + SafeWindowHwnd(activeWindow));
                    if (workbookWindow != null)
                    {
                        _logger?.Info(
                            KernelFlickerTracePrefix
                            + " source=TaskPaneRefreshOrchestrationService action=resolve-window-success reason="
                            + (reason ?? string.Empty)
                            + ", workbook="
                            + _formatWorkbookDescriptor(workbook)
                            + ", resolvedWindow="
                            + _formatWindowDescriptor(workbookWindow)
                            + ", resolveAttempt="
                            + (attempt + 1).ToString(CultureInfo.InvariantCulture));
                        return workbookWindow;
                    }

                    if (activeWorkbookMatches && activeWindow != null)
                    {
                        _logger?.Info(
                            KernelFlickerTracePrefix
                            + " source=TaskPaneRefreshOrchestrationService action=resolve-window-success-active-window reason="
                            + (reason ?? string.Empty)
                            + ", workbook="
                            + _formatWorkbookDescriptor(workbook)
                            + ", resolvedWindow="
                            + _formatWindowDescriptor(activeWindow)
                            + ", resolveAttempt="
                            + (attempt + 1).ToString(CultureInfo.InvariantCulture));
                        return activeWindow;
                    }

                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneRefreshOrchestrationService action=resolve-window-retry reason="
                        + (reason ?? string.Empty)
                        + ", workbook="
                        + _formatWorkbookDescriptor(workbook)
                        + ", resolveAttempt="
                        + (attempt + 1).ToString(CultureInfo.InvariantCulture)
                        + ", deferredToRetryCoordinator=true");
                }

                _logger?.Warn(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=resolve-window-failed reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _formatActiveState());
                _logger?.Warn(
                    "ResolveWorkbookPaneWindow failed. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + workbookFullName);
                return null;
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string FormatActiveState()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            Excel.Window activeWindow = _excelInteropService == null ? null : _excelInteropService.GetActiveWindow();
            return "activeWorkbook=" + FormatWorkbookDescriptor(activeWorkbook) + ",activeWindow=" + FormatWindowDescriptor(activeWindow);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private string SafeWorkbookName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookName(workbook);
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\"";
        }

        private static string SafeWindowCaption(Excel.Window window)
        {
            try
            {
                if (window == null)
                {
                    return string.Empty;
                }

                dynamic lateBoundWindow = window;
                return Convert.ToString(lateBoundWindow.Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    internal sealed class PendingPaneRefreshRetryService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly ExcelInteropService _excelInteropService;
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly Logger _logger;
        private readonly int _pendingPaneRefreshIntervalMs;
        private readonly int _pendingPaneRefreshMaxAttempts;
        private readonly Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> _tryRefreshTaskPane;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Action _stopPendingPaneRefreshTimer;
        private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
        private readonly Func<Excel.Window, string> _formatWindowDescriptor;
        private readonly Func<string> _formatActiveState;
        private readonly Func<Excel.Workbook, string> _safeWorkbookFullName;
        private readonly Func<Excel.Window, string> _safeWindowHwnd;
        private readonly PendingPaneRefreshRetryState _retryState = new PendingPaneRefreshRetryState();

        private System.Windows.Forms.Timer _pendingPaneRefreshTimer;

        internal PendingPaneRefreshRetryService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            int pendingPaneRefreshIntervalMs,
            int pendingPaneRefreshMaxAttempts,
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
            EnsurePendingPaneRefreshTimer();
            _pendingPaneRefreshTimer.Stop();
            _pendingPaneRefreshTimer.Start();
            return _retryState.AttemptsRemaining;
        }

        internal void StopTimer()
        {
            _pendingPaneRefreshTimer?.Stop();
        }

        private void EnsurePendingPaneRefreshTimer()
        {
            if (_pendingPaneRefreshTimer != null)
            {
                return;
            }

            _pendingPaneRefreshTimer = new System.Windows.Forms.Timer();
            _pendingPaneRefreshTimer.Interval = _pendingPaneRefreshIntervalMs;
            _pendingPaneRefreshTimer.Tick += PendingPaneRefreshTimer_Tick;
        }

        private void PendingPaneRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (!_retryState.HasAttemptsRemaining)
            {
                _stopPendingPaneRefreshTimer();
                return;
            }

            Excel.Workbook targetWorkbook = ResolvePendingPaneRefreshWorkbook();
            if (targetWorkbook != null)
            {
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
                if (refreshed)
                {
                    _stopPendingPaneRefreshTimer();
                }

                return;
            }

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
                _stopPendingPaneRefreshTimer();
                return;
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
            if (fallbackRefreshed)
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
