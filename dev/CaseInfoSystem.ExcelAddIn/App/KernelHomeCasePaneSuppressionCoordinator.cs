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

            if (string.Equals(eventName, "WorkbookOpen", StringComparison.OrdinalIgnoreCase))
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

            if (string.Equals(eventName, "WorkbookActivate", StringComparison.OrdinalIgnoreCase))
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

            if (string.Equals(eventName, "WindowActivate", StringComparison.OrdinalIgnoreCase))
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

            if (string.Equals(eventName, "WorkbookActivate", StringComparison.OrdinalIgnoreCase)
                && _suppressCasePaneOnWorkbookActivateCount > 0)
            {
                _suppressCasePaneOnWorkbookActivateCount--;
            }
            else if (string.Equals(eventName, "WindowActivate", StringComparison.OrdinalIgnoreCase)
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

            if (string.Equals(eventName, "WorkbookActivate", StringComparison.OrdinalIgnoreCase))
            {
                return _suppressCasePaneOnWorkbookActivateCount > 0;
            }

            if (string.Equals(eventName, "WindowActivate", StringComparison.OrdinalIgnoreCase))
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
    }
}
