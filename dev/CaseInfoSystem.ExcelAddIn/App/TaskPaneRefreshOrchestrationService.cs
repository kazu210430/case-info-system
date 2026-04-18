using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshOrchestrationService
    {
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
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;

        private System.Windows.Forms.Timer _pendingPaneRefreshTimer;
        private int _pendingPaneRefreshAttemptsRemaining;
        private string _pendingPaneRefreshReason;
        private string _pendingPaneRefreshWorkbookFullName;

        internal TaskPaneRefreshOrchestrationService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            TaskPaneDisplayRetryCoordinator taskPaneDisplayRetryCoordinator,
            WorkbookTaskPaneDisplayAttemptCoordinator workbookTaskPaneDisplayAttemptCoordinator,
            TaskPaneRefreshCoordinator taskPaneRefreshCoordinator,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount)
        {
            _excelInteropService = excelInteropService;
            _workbookSessionService = workbookSessionService;
            _logger = logger;
            _taskPaneDisplayRetryCoordinator = taskPaneDisplayRetryCoordinator;
            _workbookTaskPaneDisplayAttemptCoordinator = workbookTaskPaneDisplayAttemptCoordinator;
            _taskPaneRefreshCoordinator = taskPaneRefreshCoordinator;
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneRefreshCoordinator.TryRefreshTaskPane(
                reason,
                workbook,
                window,
                _getKernelHomeForm == null ? null : _getKernelHomeForm(),
                _getTaskPaneRefreshSuppressionCount == null ? 0 : _getTaskPaneRefreshSuppressionCount());
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
            _pendingPaneRefreshWorkbookFullName = string.Empty;
            if (IsTaskPaneRefreshSucceeded(reason, null, null))
            {
                StopPendingPaneRefreshTimer();
                return;
            }

            _pendingPaneRefreshReason = reason ?? string.Empty;
            _pendingPaneRefreshAttemptsRemaining = PendingPaneRefreshMaxAttempts;
            EnsurePendingPaneRefreshTimer();
            _pendingPaneRefreshTimer.Stop();
            _pendingPaneRefreshTimer.Start();
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
            _pendingPaneRefreshWorkbookFullName = _excelInteropService == null
                ? string.Empty
                : _excelInteropService.GetWorkbookFullName(workbook);
            Excel.Window workbookWindow = ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook: false);
            _logger?.Info("TaskPane timer fallback prepare. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", resolvedWindowHwnd=" + SafeWindowHwnd(workbookWindow));

            if (TryRefreshTaskPane(reason, workbook, workbookWindow).IsRefreshSucceeded)
            {
                _logger?.Info("TaskPane timer fallback immediate refresh succeeded. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook));
                StopPendingPaneRefreshTimer();
                return;
            }

            _pendingPaneRefreshReason = reason ?? string.Empty;
            _pendingPaneRefreshAttemptsRemaining = PendingPaneRefreshMaxAttempts;
            EnsurePendingPaneRefreshTimer();
            _pendingPaneRefreshTimer.Stop();
            _pendingPaneRefreshTimer.Start();
            _logger?.Info("TaskPane timer fallback scheduled. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempts=" + _pendingPaneRefreshAttemptsRemaining.ToString(CultureInfo.InvariantCulture));
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
                WaitForTaskPaneReadyRetry,
                StopPendingPaneRefreshTimer,
                ScheduleWorkbookTaskPaneRefresh);
        }

        internal Excel.Window ResolveWorkbookPaneWindow(Excel.Workbook workbook, string reason, bool activateWorkbook)
        {
            if (_excelInteropService == null || workbook == null)
            {
                return null;
            }

            string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
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
                _logger?.Info("ResolveWorkbookPaneWindow state. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", resolveAttempt=" + (attempt + 1).ToString(CultureInfo.InvariantCulture) + ", activateWorkbook=" + activateWorkbook.ToString() + ", visibleWindowHwnd=" + SafeWindowHwnd(workbookWindow) + ", activeWorkbook=" + activeWorkbookFullName + ", activeWorkbookMatches=" + string.Equals(activeWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase).ToString() + ", activeWindowHwnd=" + SafeWindowHwnd(activeWindow));
                if (workbookWindow != null)
                {
                    return workbookWindow;
                }

                if (string.Equals(activeWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase) && activeWindow != null)
                {
                    return activeWindow;
                }

                Application.DoEvents();
                Thread.Sleep(WorkbookPaneWindowResolveDelayMs);
            }

            _logger?.Warn(
                "ResolveWorkbookPaneWindow failed. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName);
            return null;
        }

        internal void StopPendingPaneRefreshTimer()
        {
            if (_pendingPaneRefreshTimer == null)
            {
                return;
            }

            _pendingPaneRefreshTimer.Stop();
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

        private bool TryShowWorkbookTaskPaneOnce(Excel.Workbook workbook, string reason, int attemptNumber)
        {
            _logger?.Info("TaskPane wait-ready attempt start. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));
            WorkbookTaskPaneDisplayAttemptResult result = _workbookTaskPaneDisplayAttemptCoordinator.TryShowOnce(
                workbook,
                reason,
                (targetWorkbook, targetReason) =>
                {
                    Excel.Window resolvedWindow = ResolveWorkbookPaneWindow(targetWorkbook, targetReason, activateWorkbook: true);
                    _logger?.Info("TaskPane wait-ready attempt window. reason=" + (targetReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture) + ", windowHwnd=" + SafeWindowHwnd(resolvedWindow) + ", activeWorkbookMatches=" + IsActiveWorkbookMatch(targetWorkbook).ToString() + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
                    return resolvedWindow;
                },
                (targetReason, targetWorkbook, targetWindow) =>
                {
                    TaskPaneRefreshAttemptResult refreshAttemptResult = TryRefreshTaskPane(targetReason, targetWorkbook, targetWindow);
                    bool attemptRefreshed = refreshAttemptResult.IsRefreshSucceeded;
                    _logger?.Info("TaskPane wait-ready attempt refresh. reason=" + (targetReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture) + ", refreshed=" + attemptRefreshed.ToString());
                    return refreshAttemptResult;
                });
            return result.RefreshAttemptResult.IsRefreshSucceeded;
        }

        private void WaitForTaskPaneReadyRetry(Excel.Workbook workbook, string reason, int attemptNumber)
        {
            _logger?.Info("TaskPane wait-ready retry wait. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempt=" + attemptNumber.ToString(CultureInfo.InvariantCulture));
            Application.DoEvents();
            Thread.Sleep(80);
        }

        private void PendingPaneRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (_pendingPaneRefreshAttemptsRemaining <= 0)
            {
                StopPendingPaneRefreshTimer();
                return;
            }

            Excel.Workbook targetWorkbook = ResolvePendingPaneRefreshWorkbook();
            if (targetWorkbook != null)
            {
                _pendingPaneRefreshAttemptsRemaining--;
                _logger?.Info("TaskPane timer retry start. reason=" + (_pendingPaneRefreshReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", attemptsRemaining=" + _pendingPaneRefreshAttemptsRemaining.ToString(CultureInfo.InvariantCulture));
                Excel.Window workbookWindow = ResolveWorkbookPaneWindow(targetWorkbook, _pendingPaneRefreshReason, activateWorkbook: true);
                bool refreshed = TryRefreshTaskPane(_pendingPaneRefreshReason, targetWorkbook, workbookWindow).IsRefreshSucceeded;
                _logger?.Info("TaskPane timer retry result. reason=" + (_pendingPaneRefreshReason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(targetWorkbook) + ", windowHwnd=" + SafeWindowHwnd(workbookWindow) + ", refreshed=" + refreshed.ToString());
                if (refreshed)
                {
                    StopPendingPaneRefreshTimer();
                }

                return;
            }

            WorkbookContext context = _workbookSessionService == null ? null : _workbookSessionService.ResolveActiveContext("PendingPaneRefresh");
            if (context == null || context.Role != WorkbookRole.Case)
            {
                StopPendingPaneRefreshTimer();
                return;
            }

            _pendingPaneRefreshAttemptsRemaining--;
            if (TryRefreshTaskPane(_pendingPaneRefreshReason, null, null).IsRefreshSucceeded)
            {
                StopPendingPaneRefreshTimer();
            }
        }

        private Excel.Workbook ResolvePendingPaneRefreshWorkbook()
        {
            if (_excelInteropService == null || string.IsNullOrWhiteSpace(_pendingPaneRefreshWorkbookFullName))
            {
                return null;
            }

            return _excelInteropService.FindOpenWorkbook(_pendingPaneRefreshWorkbookFullName);
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
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
    }
}
