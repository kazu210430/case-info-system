using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelHomeCasePaneSuppressionCoordinator
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private static readonly TimeSpan SuppressionDuration = TimeSpan.FromSeconds(5);

        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly IExcelInteropService _excelInteropService;
        private readonly Logger _logger;
        private int _suppressKernelHomeOnOpenCount;
        private int _suppressKernelHomeOnActivateCount;
        private DateTime _suppressKernelHomeUntilUtc;
        private string _suppressCasePaneWorkbookFullName;
        private int _suppressCasePaneOnWorkbookActivateCount;
        private int _suppressCasePaneOnWindowActivateCount;
        private DateTime _suppressCasePaneUntilUtc;
        private string _protectedCaseWorkbookFullName;
        private string _protectedCaseWindowHwnd;
        private DateTime _protectedCaseWorkbookActivateUntilUtc;
        private bool _kernelHomeExternalCloseBusy;
        private bool _kernelHomeExternalCloseRequested;

        internal KernelHomeCasePaneSuppressionCoordinator(
            WorkbookRoleResolver workbookRoleResolver,
            IExcelInteropService excelInteropService,
            Logger logger)
        {
            _workbookRoleResolver = workbookRoleResolver;
            _excelInteropService = excelInteropService;
            _logger = logger;
        }

        internal void SuppressUpcomingKernelHomeDisplay(string reason, bool suppressOnOpen, bool suppressOnActivate)
        {
            _suppressKernelHomeUntilUtc = DateTime.UtcNow.Add(SuppressionDuration);
            if (suppressOnOpen)
            {
                _suppressKernelHomeOnOpenCount++;
            }

            if (suppressOnActivate)
            {
                _suppressKernelHomeOnActivateCount++;
            }

            _logger?.Info(
                "[Suppression:Request] open=" + suppressOnOpen
                + ", activate=" + suppressOnActivate
                + ", until=" + _suppressKernelHomeUntilUtc.ToString("o"));
            _logger?.Info(
                "Kernel home suppression prepared. reason="
                + (reason ?? string.Empty)
                + ", suppressOnOpen="
                + suppressOnOpen.ToString()
                + ", suppressOnActivate="
                + suppressOnActivate.ToString()
                + ", suppressUntilUtc="
                + _suppressKernelHomeUntilUtc.ToString("O", CultureInfo.InvariantCulture));
        }

        internal bool ShouldSuppressKernelHomeDisplay(string eventName)
        {
            return IsKernelHomeSuppressionActive(eventName, consume: true);
        }

        internal bool IsKernelHomeSuppressionActive(string eventName, bool consume)
        {
            if (DateTime.UtcNow > _suppressKernelHomeUntilUtc)
            {
                return false;
            }

            if (string.Equals(eventName, ControlFlowReasons.WorkbookOpen, StringComparison.OrdinalIgnoreCase))
            {
                if (_suppressKernelHomeOnOpenCount > 0)
                {
                    _logger?.Info(
                        "[Suppression:State] event=" + eventName
                        + ", consume=" + consume
                        + ", openCount=" + _suppressKernelHomeOnOpenCount
                        + ", activateCount=" + _suppressKernelHomeOnActivateCount);
                    if (consume)
                    {
                        _suppressKernelHomeOnOpenCount--;
                    }

                    return true;
                }

                return false;
            }

            if (string.Equals(eventName, ControlFlowReasons.WorkbookActivate, StringComparison.OrdinalIgnoreCase))
            {
                if (_suppressKernelHomeOnActivateCount > 0)
                {
                    _logger?.Info(
                        "[Suppression:State] event=" + eventName
                        + ", consume=" + consume
                        + ", openCount=" + _suppressKernelHomeOnOpenCount
                        + ", activateCount=" + _suppressKernelHomeOnActivateCount);
                    if (consume)
                    {
                        _suppressKernelHomeOnActivateCount--;
                    }

                    return true;
                }

                return false;
            }

            if (string.Equals(eventName, ControlFlowReasons.WindowActivate, StringComparison.OrdinalIgnoreCase))
            {
                if (_suppressKernelHomeOnActivateCount > 0)
                {
                    _logger?.Info(
                        "[Suppression:State] event=" + eventName
                        + ", consume=" + consume
                        + ", openCount=" + _suppressKernelHomeOnOpenCount
                        + ", activateCount=" + _suppressKernelHomeOnActivateCount);
                    return true;
                }

                return false;
            }

            return false;
        }

        internal void HandleExternalWorkbookDetected(
            ExternalWorkbookDetectionService externalWorkbookDetectionService,
            Excel.Workbook workbook,
            string eventName,
            KernelHomeForm kernelHomeForm)
        {
            if (externalWorkbookDetectionService == null)
            {
                return;
            }

            externalWorkbookDetectionService.Handle(
                workbook,
                eventName,
                kernelHomeForm,
                IsKernelHomeSuppressionActive,
                ref _kernelHomeExternalCloseBusy,
                ref _kernelHomeExternalCloseRequested);
        }

        internal void ResetKernelHomeExternalCloseRequested()
        {
            _kernelHomeExternalCloseRequested = false;
        }

        internal void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason)
        {
            if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                return;
            }

            _suppressCasePaneWorkbookFullName = workbookFullName;
            _suppressCasePaneUntilUtc = DateTime.UtcNow.Add(SuppressionDuration);
            _suppressCasePaneOnWorkbookActivateCount = 1;
            _suppressCasePaneOnWindowActivateCount = 1;
            _logger?.Info(
                "Case pane activation suppression prepared. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName
                + ", suppressUntilUtc="
                + _suppressCasePaneUntilUtc.ToString("O", CultureInfo.InvariantCulture));
        }

        internal void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            string workbookFullName = _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
            string windowHwnd = SafeWindowHwnd(window);
            WorkbookRole workbookRole = _workbookRoleResolver == null ? WorkbookRole.Unknown : _workbookRoleResolver.Resolve(workbook);
            bool protectionStartRequested = workbookRole == WorkbookRole.Case
                && !string.IsNullOrWhiteSpace(workbookFullName)
                && !string.IsNullOrWhiteSpace(windowHwnd);
            string protectionSkipReason = string.Empty;
            if (workbookRole != WorkbookRole.Case)
            {
                protectionSkipReason = "workbookRole!=Case";
            }
            else if (string.IsNullOrWhiteSpace(windowHwnd))
            {
                protectionSkipReason = "windowHwnd=empty";
            }
            else if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                protectionSkipReason = "workbookFullName=empty";
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookActivateProtection action=evaluate reason="
                + (reason ?? string.Empty)
                + ", workbookRole="
                + workbookRole.ToString()
                + ", workbookPresent="
                + (workbook != null).ToString()
                + ", workbookFullNamePresent="
                + (!string.IsNullOrWhiteSpace(workbookFullName)).ToString()
                + ", protectedWindowPresent="
                + (!string.IsNullOrWhiteSpace(windowHwnd)).ToString()
                + ", protectionStartRequested="
                + protectionStartRequested.ToString()
                + ", protectionSkipped="
                + (!protectionStartRequested).ToString()
                + ", protectionSkipReason="
                + protectionSkipReason);
            if (!protectionStartRequested)
            {
                return;
            }

            _protectedCaseWorkbookFullName = workbookFullName;
            _protectedCaseWindowHwnd = windowHwnd;
            _protectedCaseWorkbookActivateUntilUtc = DateTime.UtcNow.Add(SuppressionDuration);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookActivateProtection action=start reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName
                + ", windowHwnd="
                + windowHwnd
                + ", protectUntilUtc="
                + _protectedCaseWorkbookActivateUntilUtc.ToString("O", CultureInfo.InvariantCulture));
        }

        internal bool ShouldIgnoreWorkbookActivateDuringProtection(Excel.Workbook workbook)
        {
            string workbookFullName = _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
            string activeWindowHwnd = SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow());
            if (!IsProtectedCaseProtectionTarget(workbookFullName, activeWindowHwnd))
            {
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookActivateProtection action=ignore event=WorkbookActivate workbook="
                + workbookFullName
                + ", activeWindowHwnd="
                + activeWindowHwnd
                + ", protectUntilUtc="
                + _protectedCaseWorkbookActivateUntilUtc.ToString("O", CultureInfo.InvariantCulture));
            return true;
        }

        internal bool ShouldIgnoreWindowActivateDuringProtection(Excel.Workbook workbook, Excel.Window window)
        {
            string workbookFullName = _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
            string windowHwnd = SafeWindowHwnd(window);
            if (!IsProtectedCaseProtectionTarget(workbookFullName, windowHwnd))
            {
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WindowActivateProtection action=ignore event=WindowActivate workbook="
                + workbookFullName
                + ", windowHwnd="
                + windowHwnd
                + ", protectUntilUtc="
                + _protectedCaseWorkbookActivateUntilUtc.ToString("O", CultureInfo.InvariantCulture));
            return true;
        }

        internal bool ShouldIgnoreTaskPaneRefreshDuringProtection(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            if (!IsProtectedCaseProtectionActive())
            {
                return false;
            }

            string activeWindowHwnd = SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow());
            if (!string.Equals(_protectedCaseWindowHwnd, activeWindowHwnd, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string workbookFullName = _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
            string windowHwnd = SafeWindowHwnd(window);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshProtection action=ignore reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName
                + ", windowHwnd="
                + windowHwnd
                + ", activeWindowHwnd="
                + activeWindowHwnd
                + ", protectUntilUtc="
                + _protectedCaseWorkbookActivateUntilUtc.ToString("O", CultureInfo.InvariantCulture));
            return true;
        }

        internal bool ShouldSuppressCasePaneRefresh(string eventName, Excel.Workbook workbook)
        {
            string workbookFullName = _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
            WorkbookRole workbookRole = _workbookRoleResolver == null ? WorkbookRole.Unknown : _workbookRoleResolver.Resolve(workbook);
            if (workbookRole == WorkbookRole.Accounting)
            {
                _logger?.Info(
                    "Accounting pane suppression probe. eventName="
                    + (eventName ?? string.Empty)
                    + ", workbook="
                    + workbookFullName
                    + ", suppressionTarget="
                    + (_suppressCasePaneWorkbookFullName ?? string.Empty)
                    + ", suppressUntilUtc="
                    + (_suppressCasePaneUntilUtc == DateTime.MinValue ? string.Empty : _suppressCasePaneUntilUtc.ToString("O", CultureInfo.InvariantCulture))
                    + ", workbookActivateRemaining="
                    + _suppressCasePaneOnWorkbookActivateCount.ToString(CultureInfo.InvariantCulture)
                    + ", windowActivateRemaining="
                    + _suppressCasePaneOnWindowActivateCount.ToString(CultureInfo.InvariantCulture));
            }

            if (!IsCasePaneSuppressionTarget(eventName, workbookFullName))
            {
                return false;
            }

            if (string.Equals(eventName, ControlFlowReasons.WorkbookActivate, StringComparison.OrdinalIgnoreCase)
                && _suppressCasePaneOnWorkbookActivateCount > 0)
            {
                _suppressCasePaneOnWorkbookActivateCount--;
            }
            else if (string.Equals(eventName, ControlFlowReasons.WindowActivate, StringComparison.OrdinalIgnoreCase)
                && _suppressCasePaneOnWindowActivateCount > 0)
            {
                _suppressCasePaneOnWindowActivateCount--;
            }

            CleanupCasePaneSuppressionIfCompleted();
            _logger?.Info(
                "Case pane refresh suppressed. eventName="
                + (eventName ?? string.Empty)
                + ", workbook="
                + workbookFullName
                + ", workbookActivateRemaining="
                + _suppressCasePaneOnWorkbookActivateCount.ToString(CultureInfo.InvariantCulture)
                + ", windowActivateRemaining="
                + _suppressCasePaneOnWindowActivateCount.ToString(CultureInfo.InvariantCulture));
            return true;
        }

        private bool IsCasePaneSuppressionTarget(string eventName, string workbookFullName)
        {
            if (DateTime.UtcNow > _suppressCasePaneUntilUtc)
            {
                ClearCasePaneSuppression("Expired");
                return false;
            }

            if (string.IsNullOrWhiteSpace(_suppressCasePaneWorkbookFullName)
                || string.IsNullOrWhiteSpace(workbookFullName)
                || !string.Equals(_suppressCasePaneWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(eventName, ControlFlowReasons.WorkbookActivate, StringComparison.OrdinalIgnoreCase))
            {
                return _suppressCasePaneOnWorkbookActivateCount > 0;
            }

            if (string.Equals(eventName, ControlFlowReasons.WindowActivate, StringComparison.OrdinalIgnoreCase))
            {
                return _suppressCasePaneOnWindowActivateCount > 0;
            }

            return false;
        }

        private void CleanupCasePaneSuppressionIfCompleted()
        {
            if (_suppressCasePaneOnWorkbookActivateCount > 0 || _suppressCasePaneOnWindowActivateCount > 0)
            {
                return;
            }

            ClearCasePaneSuppression("Consumed");
        }

        private void ClearCasePaneSuppression(string reason)
        {
            if (string.IsNullOrWhiteSpace(_suppressCasePaneWorkbookFullName)
                && _suppressCasePaneOnWorkbookActivateCount == 0
                && _suppressCasePaneOnWindowActivateCount == 0)
            {
                return;
            }

            string workbookFullName = _suppressCasePaneWorkbookFullName ?? string.Empty;
            _suppressCasePaneWorkbookFullName = string.Empty;
            _suppressCasePaneOnWorkbookActivateCount = 0;
            _suppressCasePaneOnWindowActivateCount = 0;
            _suppressCasePaneUntilUtc = DateTime.MinValue;
            _logger?.Info(
                "Case pane activation suppression cleared. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName);
        }

        private bool IsProtectedCaseProtectionTarget(string workbookFullName, string windowHwnd)
        {
            if (!IsProtectedCaseProtectionActive())
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(workbookFullName)
                || string.IsNullOrWhiteSpace(windowHwnd))
            {
                return false;
            }

            return string.Equals(_protectedCaseWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(_protectedCaseWindowHwnd, windowHwnd, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsProtectedCaseProtectionActive()
        {
            if (DateTime.UtcNow > _protectedCaseWorkbookActivateUntilUtc)
            {
                ClearCaseWorkbookActivateProtection("Expired");
                return false;
            }

            if (string.IsNullOrWhiteSpace(_protectedCaseWorkbookFullName)
                || string.IsNullOrWhiteSpace(_protectedCaseWindowHwnd))
            {
                return false;
            }

            return true;
        }

        private void ClearCaseWorkbookActivateProtection(string reason)
        {
            if (string.IsNullOrWhiteSpace(_protectedCaseWorkbookFullName)
                && string.IsNullOrWhiteSpace(_protectedCaseWindowHwnd)
                && _protectedCaseWorkbookActivateUntilUtc == DateTime.MinValue)
            {
                return;
            }

            string workbookFullName = _protectedCaseWorkbookFullName ?? string.Empty;
            string windowHwnd = _protectedCaseWindowHwnd ?? string.Empty;
            _protectedCaseWorkbookFullName = string.Empty;
            _protectedCaseWindowHwnd = string.Empty;
            _protectedCaseWorkbookActivateUntilUtc = DateTime.MinValue;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookActivateProtection action=clear reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName
                + ", windowHwnd="
                + windowHwnd);
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
