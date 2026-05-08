using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookTaskPaneReadyShowAttemptWorker
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private const int ReadyShowMaxAttempts = 2;
        private const int ReadyShowRetryDelayMs = 80;

        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;
        private readonly TaskPaneDisplayRetryCoordinator _taskPaneDisplayRetryCoordinator;
        private readonly WorkbookTaskPaneDisplayAttemptCoordinator _workbookTaskPaneDisplayAttemptCoordinator;
        private readonly WorkbookWindowVisibilityService _workbookWindowVisibilityService;
        private readonly Func<Excel.Workbook, Excel.Window, bool> _hasVisibleCasePaneForWorkbookWindow;
        private readonly Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> _tryRefreshTaskPane;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;

        internal WorkbookTaskPaneReadyShowAttemptWorker(
            ExcelInteropService excelInteropService,
            Logger logger,
            TaskPaneDisplayRetryCoordinator taskPaneDisplayRetryCoordinator,
            WorkbookTaskPaneDisplayAttemptCoordinator workbookTaskPaneDisplayAttemptCoordinator,
            WorkbookWindowVisibilityService workbookWindowVisibilityService,
            Func<Excel.Workbook, Excel.Window, bool> hasVisibleCasePaneForWorkbookWindow,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow)
        {
            _excelInteropService = excelInteropService;
            _logger = logger;
            _taskPaneDisplayRetryCoordinator = taskPaneDisplayRetryCoordinator ?? throw new ArgumentNullException(nameof(taskPaneDisplayRetryCoordinator));
            _workbookTaskPaneDisplayAttemptCoordinator = workbookTaskPaneDisplayAttemptCoordinator ?? throw new ArgumentNullException(nameof(workbookTaskPaneDisplayAttemptCoordinator));
            _workbookWindowVisibilityService = workbookWindowVisibilityService ?? throw new ArgumentNullException(nameof(workbookWindowVisibilityService));
            _hasVisibleCasePaneForWorkbookWindow = hasVisibleCasePaneForWorkbookWindow ?? throw new ArgumentNullException(nameof(hasVisibleCasePaneForWorkbookWindow));
            _tryRefreshTaskPane = tryRefreshTaskPane ?? throw new ArgumentNullException(nameof(tryRefreshTaskPane));
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow ?? throw new ArgumentNullException(nameof(resolveWorkbookPaneWindow));
        }

        internal void ShowWhenReady(
            Excel.Workbook workbook,
            string reason,
            Action<Excel.Workbook, string, int, Action> scheduleRetry,
            Action<WorkbookTaskPaneReadyShowAttemptOutcome> onShown,
            Action<Excel.Workbook, string> scheduleFallback)
        {
            if (workbook == null)
            {
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookTaskPaneReadyShowAttemptWorker action=wait-ready-entry reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", maxAttempts="
                + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryDelayMs="
                + ReadyShowRetryDelayMs.ToString(CultureInfo.InvariantCulture)
                + ", activeWorkbook="
                + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook())
                + ", activeWindowHwnd="
                + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow())
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            _logger?.Info(
                "TaskPane wait-ready start. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", maxAttempts="
                + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryDelayMs="
                + ReadyShowRetryDelayMs.ToString(CultureInfo.InvariantCulture)
                + ", activeWorkbook="
                + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook())
                + ", activeWindowHwnd="
                + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow())
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            WorkbookTaskPaneReadyShowAttemptOutcome shownOutcome = null;
            WorkbookWindowVisibilityEnsureFacts lastWorkbookWindowEnsureFacts = null;
            _taskPaneDisplayRetryCoordinator.ShowWhenReady(
                workbook,
                reason,
                (targetWorkbook, targetReason, attemptNumber) =>
                {
                    WorkbookTaskPaneReadyShowAttemptOutcome attemptOutcome = TryShowWorkbookTaskPaneOnce(targetWorkbook, targetReason, attemptNumber);
                    if (attemptOutcome.WorkbookWindowEnsureFacts != null)
                    {
                        lastWorkbookWindowEnsureFacts = attemptOutcome.WorkbookWindowEnsureFacts;
                    }

                    if (attemptOutcome.IsShown)
                    {
                        shownOutcome = attemptOutcome.WorkbookWindowEnsureFacts == null && lastWorkbookWindowEnsureFacts != null
                            ? attemptOutcome.WithWorkbookWindowEnsureFacts(lastWorkbookWindowEnsureFacts)
                            : attemptOutcome;
                    }

                    return attemptOutcome.IsShown;
                },
                scheduleRetry,
                () => onShown?.Invoke(shownOutcome),
                scheduleFallback);
        }

        private WorkbookTaskPaneReadyShowAttemptOutcome TryShowWorkbookTaskPaneOnce(Excel.Workbook workbook, string reason, int attemptNumber)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookTaskPaneReadyShowAttemptWorker action=wait-ready-attempt-start reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            _logger?.Info(
                "TaskPane wait-ready attempt start. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                null,
                "ready-show-attempt",
                "WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce",
                SafeWorkbookFullName(workbook),
                "reason=" + (reason ?? string.Empty) + ",attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));
            bool visibleCasePaneAlreadyShown = false;
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts = null;
            WorkbookTaskPaneDisplayAttemptResult result = _workbookTaskPaneDisplayAttemptCoordinator.TryShowOnce(
                workbook,
                reason,
                (targetWorkbook, targetReason) =>
                {
                    workbookWindowEnsureFacts = EnsureWorkbookWindowVisibleForTaskPaneDisplay(targetWorkbook, targetReason, attemptNumber);
                    Excel.Window resolvedWindow = _resolveWorkbookPaneWindow(targetWorkbook, targetReason, true);
                    visibleCasePaneAlreadyShown = resolvedWindow != null
                        && _hasVisibleCasePaneForWorkbookWindow(targetWorkbook, resolvedWindow);
                    if (visibleCasePaneAlreadyShown)
                    {
                        _logger?.Info(
                            "TaskPane wait-ready early-complete because visible CASE pane is already shown. reason="
                            + (targetReason ?? string.Empty)
                            + ", workbook="
                            + SafeWorkbookFullName(targetWorkbook)
                            + ", readyShowReason="
                            + (targetReason ?? string.Empty)
                            + ", attempt="
                            + attemptNumber.ToString(CultureInfo.InvariantCulture)
                            + ", maxAttempts="
                            + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                            + ", windowHwnd="
                            + SafeWindowHwnd(resolvedWindow)
                            + ", windowResolved=true"
                            + ", visibleCasePaneEarlyComplete=true"
                            + ", renderCurrentCheckBypassed=true"
                            + ", earlyCompleteBasis=retainedHost+metadataJoin+visibilityRetention");
                    }

                    _logger?.Info(
                        "TaskPane wait-ready attempt window. reason="
                        + (targetReason ?? string.Empty)
                        + ", workbook="
                        + SafeWorkbookFullName(targetWorkbook)
                        + ", attempt="
                        + attemptNumber.ToString(CultureInfo.InvariantCulture)
                        + ", maxAttempts="
                        + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                        + ", windowHwnd="
                        + SafeWindowHwnd(resolvedWindow)
                        + ", windowResolved="
                        + (resolvedWindow != null).ToString()
                        + ", visibleCasePaneEarlyComplete="
                        + visibleCasePaneAlreadyShown.ToString()
                        + ", activeWorkbookMatches="
                        + IsActiveWorkbookMatch(targetWorkbook).ToString()
                        + ", activeWindowHwnd="
                        + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
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
                        _logger?.Info(
                            "TaskPane wait-ready attempt refresh skipped because visible CASE pane is already shown. reason="
                            + (targetReason ?? string.Empty)
                            + ", workbook="
                            + SafeWorkbookFullName(targetWorkbook)
                            + ", readyShowReason="
                            + (targetReason ?? string.Empty)
                            + ", attempt="
                            + attemptNumber.ToString(CultureInfo.InvariantCulture)
                            + ", maxAttempts="
                            + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                            + ", windowHwnd="
                            + SafeWindowHwnd(targetWindow)
                            + ", windowResolved="
                            + (targetWindow != null).ToString()
                            + ", visibleCasePaneEarlyComplete=true"
                            + ", renderCurrentCheckBypassed=true"
                            + ", earlyCompleteBasis=retainedHost+metadataJoin+visibilityRetention"
                            + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, targetWorkbook));
                        NewCaseVisibilityObservation.Log(
                            _logger,
                            _excelInteropService,
                            null,
                            targetWorkbook,
                            targetWindow,
                            "taskpane-already-visible",
                            "WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce",
                            SafeWorkbookFullName(targetWorkbook),
                            "reason=" + (targetReason ?? string.Empty) + ",attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));
                        return TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied();
                    }

                    TaskPaneRefreshAttemptResult refreshAttemptResult = _tryRefreshTaskPane(targetReason, targetWorkbook, targetWindow);
                    bool attemptRefreshed = refreshAttemptResult.IsRefreshSucceeded;
                    _logger?.Info(
                        "TaskPane wait-ready attempt refresh. reason="
                        + (targetReason ?? string.Empty)
                        + ", workbook="
                        + SafeWorkbookFullName(targetWorkbook)
                        + ", attempt="
                        + attemptNumber.ToString(CultureInfo.InvariantCulture)
                        + ", maxAttempts="
                        + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                        + ", windowResolved="
                        + (targetWindow != null).ToString()
                        + ", visibleCasePaneEarlyComplete="
                        + visibleCasePaneAlreadyShown.ToString()
                        + ", refreshed="
                        + attemptRefreshed.ToString()
                        + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, targetWorkbook));
                    return refreshAttemptResult;
                });
            bool refreshed = result.RefreshAttemptResult.IsRefreshSucceeded;
            bool windowResolved = result.WorkbookWindow != null;
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                result.WorkbookWindow,
                "ready-show-attempt-result",
                "WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce",
                SafeWorkbookFullName(workbook),
                "reason=" + (reason ?? string.Empty)
                + ",attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ",windowResolved="
                + windowResolved.ToString()
                + ",refreshed="
                + refreshed.ToString()
                + ",visibleCasePaneEarlyComplete="
                + visibleCasePaneAlreadyShown.ToString());
            if (!refreshed
                && attemptNumber >= ReadyShowMaxAttempts)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=WorkbookTaskPaneReadyShowAttemptWorker action=wait-ready-attempts-exhausted reason="
                    + (reason ?? string.Empty)
                    + ", readyShowReason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + SafeWorkbookFullName(workbook)
                    + ", attempt="
                    + attemptNumber.ToString(CultureInfo.InvariantCulture)
                    + ", maxAttempts="
                    + ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                    + ", windowResolved="
                    + windowResolved.ToString()
                    + ", windowHwnd="
                    + SafeWindowHwnd(result.WorkbookWindow)
                    + ", visibleCasePaneEarlyComplete="
                    + visibleCasePaneAlreadyShown.ToString()
                    + ", fallbackCause=AttemptsExhausted"
                    + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
                NewCaseVisibilityObservation.Log(
                    _logger,
                    _excelInteropService,
                    null,
                    workbook,
                    result.WorkbookWindow,
                    "ready-show-attempts-exhausted",
                    "WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce",
                    SafeWorkbookFullName(workbook),
                    "reason=" + (reason ?? string.Empty)
                    + ",attempt="
                    + attemptNumber.ToString(CultureInfo.InvariantCulture)
                    + ",windowResolved="
                    + windowResolved.ToString());
            }

            return new WorkbookTaskPaneReadyShowAttemptOutcome(
                attemptNumber,
                result.WorkbookWindow,
                result.RefreshAttemptResult,
                visibleCasePaneAlreadyShown,
                workbookWindowEnsureFacts);
        }

        private WorkbookWindowVisibilityEnsureFacts EnsureWorkbookWindowVisibleForTaskPaneDisplay(Excel.Workbook workbook, string reason, int attemptNumber)
        {
            if (attemptNumber != 1 || workbook == null)
            {
                return null;
            }

            WorkbookWindowVisibilityEnsureResult result = _workbookWindowVisibilityService.EnsureVisible(workbook, reason);
            _logger?.Info(
                "TaskPane wait-ready pre-visibility evaluated. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + result.WorkbookFullName
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", outcome="
                + result.Outcome.ToString()
                + ", windowHwnd="
                + result.WindowHwnd
                + ", visibleAfterSet="
                + (result.VisibleAfterSet.HasValue ? result.VisibleAfterSet.Value.ToString() : string.Empty));
            return WorkbookWindowVisibilityEnsureFacts.FromResult(result);
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

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
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
    }
}
