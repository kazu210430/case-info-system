using System;
using System.Diagnostics;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class ShutdownCleanupAdapter
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Logger _logger;
        private readonly Excel.Application _application;
        private readonly Action _unhookApplicationEvents;
        private readonly Action _stopPendingPaneRefreshTimer;
        private readonly Action _closeKernelHomeOnShutdown;
        private readonly Action _disposeTaskPanes;
        private readonly Action _stopWordWarmupTimer;
        private readonly Action _stopManagedCloseStartupGuardTimer;
        private readonly Action _shutdownHiddenApplicationCache;
        private readonly Func<bool> _customTaskPanesPresent;
        private readonly Func<int> _customTaskPanesCount;

        internal ShutdownCleanupAdapter(
            Logger logger,
            Excel.Application application,
            Action unhookApplicationEvents,
            Action stopPendingPaneRefreshTimer,
            Action closeKernelHomeOnShutdown,
            Action disposeTaskPanes,
            Action stopWordWarmupTimer,
            Action stopManagedCloseStartupGuardTimer,
            Action shutdownHiddenApplicationCache,
            Func<bool> customTaskPanesPresent,
            Func<int> customTaskPanesCount)
        {
            _logger = logger;
            _application = application;
            _unhookApplicationEvents = unhookApplicationEvents;
            _stopPendingPaneRefreshTimer = stopPendingPaneRefreshTimer;
            _closeKernelHomeOnShutdown = closeKernelHomeOnShutdown;
            _disposeTaskPanes = disposeTaskPanes;
            _stopWordWarmupTimer = stopWordWarmupTimer;
            _stopManagedCloseStartupGuardTimer = stopManagedCloseStartupGuardTimer;
            _shutdownHiddenApplicationCache = shutdownHiddenApplicationCache;
            _customTaskPanesPresent = customTaskPanesPresent;
            _customTaskPanesCount = customTaskPanesCount;
        }

        internal void HandleShutdown()
        {
            if (_logger != null)
            {
                _logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-entry");
                LogShutdownState("handler-entry");
                RunShutdownStep("UnhookApplicationEvents", _unhookApplicationEvents);
                RunShutdownStep("StopPendingPaneRefreshTimer", _stopPendingPaneRefreshTimer);
                RunShutdownStep("KernelHomeFormHost.CloseOnShutdown", _closeKernelHomeOnShutdown);
                RunShutdownStep(
                    "TaskPaneManager.DisposeAll",
                    () =>
                    {
                        LogShutdownState("before-taskpane-manager-disposeall");
                        _disposeTaskPanes?.Invoke();
                        LogShutdownState("after-taskpane-manager-disposeall");
                    });
                RunShutdownStep("StopWordWarmupTimer", _stopWordWarmupTimer);
                RunShutdownStep("StopManagedCloseStartupGuardTimer", _stopManagedCloseStartupGuardTimer);
                RunShutdownStep("ShutdownHiddenApplicationCache", _shutdownHiddenApplicationCache);

                LogShutdownState("before-generated-base-boundary");
                _logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-before-generated-base-boundary");
                _logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-exit");
                _logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-complete");
                _logger.Info("ThisAddIn_Shutdown fired.");
                return;
            }

            WriteFallbackShutdownTrace();
        }

        internal static void WriteFallbackShutdownTrace()
        {
            ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-entry logger=null");
            ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-before-generated-base-boundary logger=null");
            ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-exit logger=null");
            ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-complete logger=null");
            ExcelAddInTraceLogWriter.Write("ThisAddIn_Shutdown fired.");
        }

        internal void TraceGeneratedOnShutdownBoundary(string phase)
        {
            SafeWriteShutdownTrace(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=generated-onshutdown-boundary phase="
                + (phase ?? string.Empty)
                + ", pid="
                + Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
                + ", "
                + CaptureCustomTaskPaneFacts(_customTaskPanesPresent, _customTaskPanesCount));
        }

        internal static void TraceGeneratedOnShutdownBoundaryFallback(
            string phase,
            Func<bool> customTaskPanesPresent,
            Func<int> customTaskPanesCount)
        {
            SafeWriteShutdownTraceFallback(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=generated-onshutdown-boundary phase="
                + (phase ?? string.Empty)
                + ", pid="
                + Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
                + ", "
                + CaptureCustomTaskPaneFacts(customTaskPanesPresent, customTaskPanesCount));
        }

        private void RunShutdownStep(string stepName, Action action)
        {
            _logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-step-start step=" + stepName);

            try
            {
                action?.Invoke();
                _logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-step-complete step=" + stepName);
            }
            catch (Exception exception)
            {
                _logger.Error(
                    "[KernelFlickerTrace] source=ThisAddIn action=shutdown-step-failure step="
                        + stepName
                        + ", exceptionType="
                        + exception.GetType().Name
                        + ", hResult=0x"
                        + exception.HResult.ToString("X8", CultureInfo.InvariantCulture)
                        + ", message="
                        + exception.Message,
                    exception);
                throw;
            }
        }

        private void LogShutdownState(string phase)
        {
            SafeWriteShutdownTrace(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=shutdown-state phase="
                + (phase ?? string.Empty)
                + ", "
                + CaptureShutdownExcelStateFacts()
                + ", "
                + CaptureCustomTaskPaneFacts());
        }

        private string CaptureShutdownExcelStateFacts()
        {
            string applicationVisible = ReadShutdownValue(() => _application.Visible, out bool applicationVisibleReadFailed);
            string workbooksCount = ReadShutdownValue(() => _application.Workbooks.Count, out bool workbooksCountReadFailed);
            string windowsCount = ReadShutdownValue(() => _application.Windows.Count, out bool windowsCountReadFailed);
            string displayAlerts = ReadShutdownValue(() => _application.DisplayAlerts, out bool displayAlertsReadFailed);
            string enableEvents = ReadShutdownValue(() => _application.EnableEvents, out bool enableEventsReadFailed);
            string calculationState = ReadShutdownValue(() => _application.CalculationState, out bool calculationStateReadFailed);
            string hwnd = ReadShutdownValue(() => _application.Hwnd, out bool hwndReadFailed);

            return "pid="
                + Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
                + ", applicationPresent="
                + (_application != null).ToString()
                + ", applicationVisible="
                + applicationVisible
                + ", applicationVisibleReadFailed="
                + applicationVisibleReadFailed.ToString()
                + ", workbooksCount="
                + workbooksCount
                + ", workbooksCountReadFailed="
                + workbooksCountReadFailed.ToString()
                + ", "
                + CaptureShutdownActiveWorkbookFacts()
                + ", windowsCount="
                + windowsCount
                + ", windowsCountReadFailed="
                + windowsCountReadFailed.ToString()
                + ", displayAlerts="
                + displayAlerts
                + ", displayAlertsReadFailed="
                + displayAlertsReadFailed.ToString()
                + ", enableEvents="
                + enableEvents
                + ", enableEventsReadFailed="
                + enableEventsReadFailed.ToString()
                + ", calculationState="
                + calculationState
                + ", calculationStateReadFailed="
                + calculationStateReadFailed.ToString()
                + ", hwnd="
                + hwnd
                + ", hwndReadFailed="
                + hwndReadFailed.ToString();
        }

        private string CaptureShutdownActiveWorkbookFacts()
        {
            bool activeWorkbookPresent = false;
            bool activeWorkbookReadFailed = false;
            bool activeWorkbookNameReadFailed = false;
            string activeWorkbookName = string.Empty;

            try
            {
                Excel.Workbook activeWorkbook = _application == null ? null : _application.ActiveWorkbook;
                activeWorkbookPresent = activeWorkbook != null;
                if (activeWorkbook != null)
                {
                    try
                    {
                        activeWorkbookName = SanitizeShutdownLogValue(activeWorkbook.Name);
                    }
                    catch (Exception)
                    {
                        activeWorkbookNameReadFailed = true;
                    }
                }
            }
            catch (Exception)
            {
                activeWorkbookReadFailed = true;
            }

            return "activeWorkbookPresent="
                + activeWorkbookPresent.ToString()
                + ", activeWorkbookName=\""
                + activeWorkbookName
                + "\", activeWorkbookReadFailed="
                + activeWorkbookReadFailed.ToString()
                + ", activeWorkbookNameReadFailed="
                + activeWorkbookNameReadFailed.ToString();
        }

        private string CaptureCustomTaskPaneFacts()
        {
            return CaptureCustomTaskPaneFacts(_customTaskPanesPresent, _customTaskPanesCount);
        }

        private static string CaptureCustomTaskPaneFacts(Func<bool> customTaskPanesPresentAccessor, Func<int> customTaskPanesCountAccessor)
        {
            string customTaskPanesCount = ReadShutdownValue(
                () => customTaskPanesCountAccessor == null ? 0 : customTaskPanesCountAccessor(),
                out bool customTaskPanesCountReadFailed);
            bool customTaskPanesPresent = false;
            try
            {
                customTaskPanesPresent = customTaskPanesPresentAccessor != null && customTaskPanesPresentAccessor();
            }
            catch (Exception)
            {
                customTaskPanesPresent = false;
            }

            return "customTaskPanesPresent="
                + customTaskPanesPresent.ToString()
                + ", customTaskPanesCount="
                + customTaskPanesCount
                + ", customTaskPanesCountReadFailed="
                + customTaskPanesCountReadFailed.ToString();
        }

        private static string ReadShutdownValue<T>(Func<T> read, out bool readFailed)
        {
            readFailed = false;
            try
            {
                T value = read();
                return SanitizeShutdownLogValue(value == null ? string.Empty : value.ToString());
            }
            catch (Exception)
            {
                readFailed = true;
                return string.Empty;
            }
        }

        private static string SanitizeShutdownLogValue(string value)
        {
            return (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
        }

        private void SafeWriteShutdownTrace(string message)
        {
            try
            {
                if (_logger != null)
                {
                    _logger.Info(message);
                    return;
                }

                ExcelAddInTraceLogWriter.Write((message ?? string.Empty) + " logger=null");
            }
            catch (Exception)
            {
                // Shutdown diagnostics must never change the unload control flow.
            }
        }

        private static void SafeWriteShutdownTraceFallback(string message)
        {
            try
            {
                ExcelAddInTraceLogWriter.Write((message ?? string.Empty) + " logger=null");
            }
            catch (Exception)
            {
                // Shutdown diagnostics must never change the unload control flow.
            }
        }
    }
}
