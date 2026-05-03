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
        private readonly TaskPaneDisplayRetryCoordinator _taskPaneDisplayRetryCoordinator;
        private readonly WorkbookTaskPaneDisplayAttemptCoordinator _workbookTaskPaneDisplayAttemptCoordinator;
        private readonly TaskPaneRefreshCoordinator _taskPaneRefreshCoordinator;
        private readonly WorkbookPaneWindowResolver _workbookPaneWindowResolver;
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly ICasePaneHostBridge _casePaneHostBridge;

        private System.Windows.Forms.Timer _pendingPaneRefreshTimer;
        private readonly List<System.Windows.Forms.Timer> _waitReadyRetryTimers = new List<System.Windows.Forms.Timer>();
        private readonly PendingPaneRefreshRetryState _pendingPaneRefreshRetryState = new PendingPaneRefreshRetryState();
        private int _kernelFlickerTraceRefreshAttemptSequence;

        internal TaskPaneRefreshOrchestrationService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            TaskPaneDisplayRetryCoordinator taskPaneDisplayRetryCoordinator,
            WorkbookTaskPaneDisplayAttemptCoordinator workbookTaskPaneDisplayAttemptCoordinator,
            TaskPaneRefreshCoordinator taskPaneRefreshCoordinator,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount,
            ICasePaneHostBridge casePaneHostBridge)
        {
            _excelInteropService = excelInteropService;
            _workbookSessionService = workbookSessionService;
            _logger = logger;
            _taskPaneDisplayRetryCoordinator = taskPaneDisplayRetryCoordinator;
            _workbookTaskPaneDisplayAttemptCoordinator = workbookTaskPaneDisplayAttemptCoordinator;
            _taskPaneRefreshCoordinator = taskPaneRefreshCoordinator;
            _workbookPaneWindowResolver = new WorkbookPaneWindowResolver(
                _excelInteropService,
                _logger,
                workbook => FormatWorkbookDescriptor(workbook),
                window => FormatWindowDescriptor(window),
                () => FormatActiveState());
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
            _casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException(nameof(casePaneHostBridge));
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
            if (TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=skip-workbook-open-window-dependent-refresh refreshAttemptId="
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

            if (_casePaneHostBridge.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=ignore-during-protection refreshAttemptId="
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

            TaskPaneRefreshAttemptResult result = _taskPaneRefreshCoordinator.TryRefreshTaskPane(
                reason,
                workbook,
                window,
                _getKernelHomeForm == null ? null : _getKernelHomeForm(),
                _getTaskPaneRefreshSuppressionCount == null ? 0 : _getTaskPaneRefreshSuppressionCount());
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
                + (result == null ? "null" : result.IsRefreshSucceeded.ToString()));
            return result;
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
            _pendingPaneRefreshRetryState.TrackActiveTarget();
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

            BeginPendingPaneRefreshRetry(reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (reason ?? string.Empty)
                + ", target=active"
                + ", attempts="
                + _pendingPaneRefreshRetryState.AttemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
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

            _pendingPaneRefreshRetryState.TrackWorkbookTarget(_excelInteropService == null
                ? string.Empty
                : _excelInteropService.GetWorkbookFullName(workbook));
            Excel.Window workbookWindow = ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook: false);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-prepare reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", resolvedWindow="
                + FormatWindowDescriptor(workbookWindow)
                + ", activeState="
                + FormatActiveState());
            _logger?.Info("TaskPane timer fallback prepare. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", resolvedWindowHwnd=" + SafeWindowHwnd(workbookWindow));

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

            BeginPendingPaneRefreshRetry(reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", attempts="
                + _pendingPaneRefreshRetryState.AttemptsRemaining.ToString(CultureInfo.InvariantCulture));
            _logger?.Info("TaskPane timer fallback scheduled. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempts=" + _pendingPaneRefreshRetryState.AttemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        internal void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            if (workbook == null)
            {
                return;
            }

            _logger?.Info("TaskPane wait-ready start. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", activeWorkbook=" + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook()) + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
            _taskPaneDisplayRetryCoordinator.ShowWhenReady(
                workbook,
                reason,
                TryShowWorkbookTaskPaneOnce,
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
            if (_pendingPaneRefreshTimer == null)
            {
                StopWaitReadyRetryTimers();
                return;
            }

            _pendingPaneRefreshTimer.Stop();
            StopWaitReadyRetryTimers();
        }

        private void EnsurePendingPaneRefreshTimer()
        {
            if (_pendingPaneRefreshTimer != null)
            {
                return;
            }

            _pendingPaneRefreshTimer = new System.Windows.Forms.Timer();
            _pendingPaneRefreshTimer.Interval = PendingPaneRefreshIntervalMs;
            _pendingPaneRefreshTimer.Tick += PendingPaneRefreshTimer_Tick;
        }

        private void BeginPendingPaneRefreshRetry(string reason)
        {
            _pendingPaneRefreshRetryState.BeginRetrySequence(reason, PendingPaneRefreshMaxAttempts);
            EnsurePendingPaneRefreshTimer();
            _pendingPaneRefreshTimer.Stop();
            _pendingPaneRefreshTimer.Start();
        }

        private bool TryShowWorkbookTaskPaneOnce(Excel.Workbook workbook, string reason, int attemptNumber)
        {
            _logger?.Info("TaskPane wait-ready attempt start. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));
            bool visibleCasePaneAlreadyShown = false;
            WorkbookTaskPaneDisplayAttemptResult result = _workbookTaskPaneDisplayAttemptCoordinator.TryShowOnce(
                workbook,
                reason,
                (targetWorkbook, targetReason) =>
                {
                    EnsureWorkbookWindowVisibleForTaskPaneDisplay(targetWorkbook, targetReason, attemptNumber);
                    Excel.Window resolvedWindow = ResolveWorkbookPaneWindow(targetWorkbook, targetReason, activateWorkbook: true);
                    visibleCasePaneAlreadyShown = resolvedWindow != null
                        && _casePaneHostBridge.HasVisibleCasePaneForWorkbookWindow(targetWorkbook, resolvedWindow);
                    if (visibleCasePaneAlreadyShown)
                    {
                        _logger?.Info("TaskPane wait-ready early-complete because visible CASE pane is already shown. reason=" + (targetReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(resolvedWindow));
                    }

                    _logger?.Info("TaskPane wait-ready attempt window. reason=" + (targetReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(resolvedWindow) + ", activeWorkbookMatches=" + IsActiveWorkbookMatch(targetWorkbook).ToString() + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
                    return resolvedWindow;
                },
                (targetReason, targetWorkbook, targetWindow) =>
                {
                    if (visibleCasePaneAlreadyShown)
                    {
                        NewCaseDefaultTimingLogHelper.LogTaskPaneReadyWaitToRefreshCompleted(
                            _logger,
                            SafeWorkbookFullName(targetWorkbook),
                            targetReason,
                            refreshed: false,
                            completion: "visibleCasePaneAlreadyShown");
                        _logger?.Info("TaskPane wait-ready attempt refresh skipped because visible CASE pane is already shown. reason=" + (targetReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(targetWindow));
                        return TaskPaneRefreshAttemptResult.Succeeded();
                    }

                    TaskPaneRefreshAttemptResult refreshAttemptResult = TryRefreshTaskPane(targetReason, targetWorkbook, targetWindow);
                    bool attemptRefreshed = refreshAttemptResult.IsRefreshSucceeded;
                    _logger?.Info("TaskPane wait-ready attempt refresh. reason=" + (targetReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture) + ", refreshed=" + attemptRefreshed.ToString());
                    return refreshAttemptResult;
                });
            return result.RefreshAttemptResult.IsRefreshSucceeded;
        }

        private void EnsureWorkbookWindowVisibleForTaskPaneDisplay(Excel.Workbook workbook, string reason, int attemptNumber)
        {
            if (attemptNumber != 1 || workbook == null)
            {
                return;
            }

            try
            {
                Stopwatch ensureVisibilityStopwatch = Stopwatch.StartNew();
                string workbookFullName = SafeWorkbookFullName(workbook);
                Excel.Window workbookWindow = null;
                int workbookWindowsCount = -1;

                if (_excelInteropService == null)
                {
                    _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=GetFirstVisibleWindow, skipped=true, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    long beforeGetFirstVisibleWindowElapsedMs = ensureVisibilityStopwatch.ElapsedMilliseconds;
                    _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=GetFirstVisibleWindow, phase=before, elapsedMs=" + beforeGetFirstVisibleWindowElapsedMs.ToString(CultureInfo.InvariantCulture));
                    workbookWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
                    _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=GetFirstVisibleWindow, phase=after, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow));
                }

                if (workbookWindow == null)
                {
                    long beforeWorkbookWindowsCountElapsedMs = ensureVisibilityStopwatch.ElapsedMilliseconds;
                    _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=WorkbookWindows.Count, phase=before, elapsedMs=" + beforeWorkbookWindowsCountElapsedMs.ToString(CultureInfo.InvariantCulture));
                    workbookWindowsCount = workbook.Windows.Count;
                    _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=WorkbookWindows.Count, phase=after, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", count=" + workbookWindowsCount.ToString(CultureInfo.InvariantCulture));

                    if (workbookWindowsCount > 0)
                    {
                        long beforeWorkbookWindowIndexElapsedMs = ensureVisibilityStopwatch.ElapsedMilliseconds;
                        _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=WorkbookWindows[1], phase=before, elapsedMs=" + beforeWorkbookWindowIndexElapsedMs.ToString(CultureInfo.InvariantCulture));
                        workbookWindow = workbook.Windows[1];
                        _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=WorkbookWindows[1], phase=after, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow));
                    }
                }

                if (workbookWindow == null)
                {
                    _logger?.Warn("TaskPane wait-ready pre-visibility skipped because workbook window could not be resolved. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                    return;
                }

                long beforeVisibleGetElapsedMs = ensureVisibilityStopwatch.ElapsedMilliseconds;
                _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=Window.Visible(get), phase=before, elapsedMs=" + beforeVisibleGetElapsedMs.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow));
                bool isVisible = workbookWindow.Visible;
                _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=Window.Visible(get), phase=after, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow) + ", visible=" + isVisible.ToString());

                _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=Window.Activate, skipped=true, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", note=not-invoked-by-this-method");
                _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=Window.WindowState, skipped=true, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", note=not-invoked-by-this-method");

                if (isVisible)
                {
                    _logger?.Info("TaskPane wait-ready pre-visibility skipped because workbook window is already visible. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", windowHwnd=" + SafeWindowHwnd(workbookWindow) + ", elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                    return;
                }

                long beforeVisibleSetElapsedMs = ensureVisibilityStopwatch.ElapsedMilliseconds;
                _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=Window.Visible(set:true), phase=before, elapsedMs=" + beforeVisibleSetElapsedMs.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow));
                workbookWindow.Visible = true;
                _logger?.Info("TaskPane wait-ready pre-visibility timing. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", step=Window.Visible(set:true), phase=after, elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow));
                _logger?.Info("TaskPane wait-ready workbook window made visible. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", windowHwnd=" + SafeWindowHwnd(workbookWindow) + ", elapsedMs=" + ensureVisibilityStopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                _logger?.Error("EnsureWorkbookWindowVisibleForTaskPaneDisplay failed. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook), ex);
            }
        }

        private void ScheduleTaskPaneReadyRetry(Excel.Workbook workbook, string reason, int attemptNumber, Action retryAction)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=wait-ready-retry-scheduled reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", delayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
            _logger?.Info("TaskPane wait-ready retry scheduled. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));

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
                _logger?.Info("TaskPane wait-ready retry firing. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));
                retryAction();
            };

            _waitReadyRetryTimers.Add(retryTimer);
            retryTimer.Tick += tickHandler;
            retryTimer.Start();
        }

        private void PendingPaneRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (!_pendingPaneRefreshRetryState.HasAttemptsRemaining)
            {
                StopPendingPaneRefreshTimer();
                return;
            }

            Excel.Workbook targetWorkbook = ResolvePendingPaneRefreshWorkbook();
            if (targetWorkbook != null)
            {
                int attemptsRemaining = _pendingPaneRefreshRetryState.ConsumeAttempt();
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-retry-start reason="
                    + _pendingPaneRefreshRetryState.Reason
                    + ", workbook="
                    + FormatWorkbookDescriptor(targetWorkbook)
                    + ", attemptsRemaining="
                    + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
                _logger?.Info("TaskPane timer retry start. reason=" + _pendingPaneRefreshRetryState.Reason + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attemptsRemaining=" + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
                Excel.Window workbookWindow = ResolveWorkbookPaneWindow(targetWorkbook, _pendingPaneRefreshRetryState.Reason, activateWorkbook: true);
                bool refreshed = TryRefreshTaskPane(_pendingPaneRefreshRetryState.Reason, targetWorkbook, workbookWindow).IsRefreshSucceeded;
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-retry-end reason="
                    + _pendingPaneRefreshRetryState.Reason
                    + ", workbook="
                    + FormatWorkbookDescriptor(targetWorkbook)
                    + ", resolvedWindow="
                    + FormatWindowDescriptor(workbookWindow)
                    + ", refreshed="
                    + refreshed.ToString());
                _logger?.Info("TaskPane timer retry result. reason=" + _pendingPaneRefreshRetryState.Reason + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow) + ", refreshed=" + refreshed.ToString());
                if (refreshed)
                {
                    StopPendingPaneRefreshTimer();
                }

                return;
            }

            WorkbookContext context = _workbookSessionService == null ? null : _workbookSessionService.ResolveActiveContext("PendingPaneRefresh");
            if (context == null || context.Role != WorkbookRole.Case)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-stop reason="
                    + _pendingPaneRefreshRetryState.Reason
                    + ", pendingWorkbookFullName=\""
                    + _pendingPaneRefreshRetryState.WorkbookFullName
                    + "\", contextRole="
                    + (context == null ? "null" : context.Role.ToString())
                    + ", attemptsRemaining="
                    + _pendingPaneRefreshRetryState.AttemptsRemaining.ToString(CultureInfo.InvariantCulture));
                StopPendingPaneRefreshTimer();
                return;
            }

            int fallbackAttemptsRemaining = _pendingPaneRefreshRetryState.ConsumeAttempt();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-start reason="
                + _pendingPaneRefreshRetryState.Reason
                + ", pendingWorkbookFullName=\""
                + _pendingPaneRefreshRetryState.WorkbookFullName
                + "\", contextWorkbook="
                + FormatWorkbookDescriptor(context.Workbook)
                + ", attemptsRemaining="
                + fallbackAttemptsRemaining.ToString(CultureInfo.InvariantCulture)
                + ", activeState="
                + FormatActiveState());
            _logger?.Info(
                "TaskPane timer fallback active CASE context start. reason="
                + _pendingPaneRefreshRetryState.Reason
                + ", pendingWorkbook="
                + _pendingPaneRefreshRetryState.WorkbookFullName
                + ", attemptsRemaining="
                + fallbackAttemptsRemaining.ToString(CultureInfo.InvariantCulture));
            bool fallbackRefreshed = TryRefreshTaskPane(_pendingPaneRefreshRetryState.Reason, null, null).IsRefreshSucceeded;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-end reason="
                + _pendingPaneRefreshRetryState.Reason
                + ", pendingWorkbookFullName=\""
                + _pendingPaneRefreshRetryState.WorkbookFullName
                + "\", contextWorkbook="
                + FormatWorkbookDescriptor(context.Workbook)
                + ", refreshed="
                + fallbackRefreshed.ToString()
                + ", activeState="
                + FormatActiveState());
            _logger?.Info(
                "TaskPane timer fallback active CASE context result. reason="
                + _pendingPaneRefreshRetryState.Reason
                + ", pendingWorkbook="
                + _pendingPaneRefreshRetryState.WorkbookFullName
                + ", refreshed="
                + fallbackRefreshed.ToString());
            if (fallbackRefreshed)
            {
                StopPendingPaneRefreshTimer();
            }
        }

        private Excel.Workbook ResolvePendingPaneRefreshWorkbook()
        {
            if (_excelInteropService == null || string.IsNullOrWhiteSpace(_pendingPaneRefreshRetryState.WorkbookFullName))
            {
                return null;
            }

            return _excelInteropService.FindOpenWorkbook(_pendingPaneRefreshRetryState.WorkbookFullName);
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

        private bool IsActiveWorkbookMatch(Excel.Workbook workbook)
        {
            if (_excelInteropService == null || workbook == null)
            {
                return false;
            }

            string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
            string activeWorkbookFullName = _excelInteropService.GetWorkbookFullName(_excelInteropService.GetActiveWorkbook());
            return string.Equals(workbookFullName, activeWorkbookFullName, StringComparison.OrdinalIgnoreCase);
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
}
