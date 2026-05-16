using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class AddInStartupBoundaryCoordinator
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Excel.Application _application;
        private readonly Logger _logger;
        private readonly ManagedWorkbookCloseMarkerStore _managedWorkbookCloseMarkerStore;
        private readonly Func<bool> _shouldShowKernelHomeOnStartup;
        private readonly Func<string> _describeStartupState;
        private readonly Action<string> _clearHomeWorkbookBinding;
        private readonly Action _showKernelHomePlaceholder;
        private readonly Action<string, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Func<Excel.Workbook> _getActiveWorkbook;
        private readonly Func<Excel.Workbook, string> _getWorkbookName;
        private readonly Func<Excel.Workbook, bool> _isKernelWorkbook;
        private Timer _managedCloseStartupGuardTimer;
        private bool _workbookOpenObservedSinceStartup;

        internal AddInStartupBoundaryCoordinator(
            Excel.Application application,
            Logger logger,
            ManagedWorkbookCloseMarkerStore managedWorkbookCloseMarkerStore,
            Func<bool> shouldShowKernelHomeOnStartup,
            Func<string> describeStartupState,
            Action<string> clearHomeWorkbookBinding,
            Action showKernelHomePlaceholder,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPane,
            Func<Excel.Workbook> getActiveWorkbook,
            Func<Excel.Workbook, string> getWorkbookName,
            Func<Excel.Workbook, bool> isKernelWorkbook)
        {
            _application = application;
            _logger = logger;
            _managedWorkbookCloseMarkerStore = managedWorkbookCloseMarkerStore;
            _shouldShowKernelHomeOnStartup = shouldShowKernelHomeOnStartup;
            _describeStartupState = describeStartupState;
            _clearHomeWorkbookBinding = clearHomeWorkbookBinding;
            _showKernelHomePlaceholder = showKernelHomePlaceholder;
            _refreshTaskPane = refreshTaskPane;
            _getActiveWorkbook = getActiveWorkbook;
            _getWorkbookName = getWorkbookName;
            _isKernelWorkbook = isKernelWorkbook;
        }

        internal void MarkWorkbookOpenObserved()
        {
            _workbookOpenObservedSinceStartup = true;
        }

        internal void RunAfterApplicationEventsHooked()
        {
            TryShowKernelHomeFormOnStartup();
            _refreshTaskPane?.Invoke("Startup", null, null);
            TraceAndScheduleManagedCloseStartupGuard();
        }

        internal void StopManagedCloseStartupGuardTimer()
        {
            if (_managedCloseStartupGuardTimer == null)
            {
                return;
            }

            _managedCloseStartupGuardTimer.Stop();
            _managedCloseStartupGuardTimer.Dispose();
            _managedCloseStartupGuardTimer = null;
        }

        private void TryShowKernelHomeFormOnStartup()
        {
            bool shouldShow = _shouldShowKernelHomeOnStartup != null && _shouldShowKernelHomeOnStartup();
            _logger.Info("TryShowKernelHomeFormOnStartup shouldShow=" + shouldShow + ", " + DescribeStartupState());
            if (!shouldShow)
            {
                return;
            }

            _clearHomeWorkbookBinding?.Invoke("ThisAddIn.TryShowKernelHomeFormOnStartup");
            _showKernelHomePlaceholder?.Invoke();
        }

        private string DescribeStartupState()
        {
            try
            {
                return _describeStartupState == null ? string.Empty : _describeStartupState();
            }
            catch
            {
                return string.Empty;
            }
        }

        internal void TraceAndScheduleManagedCloseStartupGuard()
        {
            ManagedWorkbookCloseMarkerReadResult markerResult = null;
            if (_managedWorkbookCloseMarkerStore != null)
            {
                try
                {
                    markerResult = _managedWorkbookCloseMarkerStore.Consume();
                }
                catch (Exception ex)
                {
                    _logger?.Error("Managed close startup marker consume failed.", ex);
                }
            }

            LogManagedCloseStartupMarker(markerResult);
            bool hasValidStartupMarker = markerResult != null && markerResult.IsValid;
            ManagedCloseStartupFacts startupFacts = CaptureManagedCloseStartupFacts("startup");
            LogManagedCloseStartupFacts(startupFacts, markerResult);
            if (!hasValidStartupMarker)
            {
                return;
            }

            ManagedCloseStartupGuardDelayDecision startupGuardDecision = ManagedCloseStartupGuardPolicy.Decide(ToManagedCloseStartupGuardFacts(startupFacts));
            if (!startupGuardDecision.IsEligible)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-guard-skip"
                    + " phase=startup"
                    + " reason=startupFactsNotEligible"
                    + FormatManagedCloseMarkerFields(markerResult)
                    + startupFacts.ToTraceFields());
                return;
            }

            StopManagedCloseStartupGuardTimer();
            _managedCloseStartupGuardTimer = new Timer();
            _managedCloseStartupGuardTimer.Interval = startupGuardDecision.DelayMs;
            _managedCloseStartupGuardTimer.Tick += (sender, args) =>
            {
                StopManagedCloseStartupGuardTimer();
                ExecuteManagedCloseStartupGuard(markerResult);
            };
            _managedCloseStartupGuardTimer.Start();
            _logger?.Info(
                KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-guard-scheduled"
                + " delayMs=" + startupGuardDecision.DelayMs.ToString(CultureInfo.InvariantCulture)
                + ", delayReason=" + startupGuardDecision.DelayReason
                + ", guardedRestoreEmptyStartupDelay=" + startupGuardDecision.UsesGuardedRestoreEmptyStartupDelay.ToString()
                + FormatManagedCloseMarkerFields(markerResult)
                + startupFacts.ToTraceFields());
        }

        internal void ExecuteManagedCloseStartupGuard(ManagedWorkbookCloseMarkerReadResult markerResult)
        {
            ManagedCloseStartupFacts delayedFacts = CaptureManagedCloseStartupFacts("delayed");
            LogManagedCloseStartupFacts(delayedFacts, markerResult);
            if (!IsManagedCloseStartupGuardEligible(delayedFacts))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-guard-skip"
                    + " phase=delayed"
                    + " reason=delayedFactsNotEligible"
                    + FormatManagedCloseMarkerFields(markerResult)
                    + delayedFacts.ToTraceFields());
                return;
            }

            ManagedCloseStartupFacts preQuitFacts = CaptureManagedCloseStartupFacts("preQuit");
            LogManagedCloseStartupFacts(preQuitFacts, markerResult);
            if (!IsManagedCloseStartupGuardEligible(preQuitFacts))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-guard-skip"
                    + " phase=preQuit"
                    + " reason=preQuitFactsNotEligible"
                    + FormatManagedCloseMarkerFields(markerResult)
                    + preQuitFacts.ToTraceFields());
                return;
            }

            QuitEmptyStartupExcelForManagedClose(markerResult, preQuitFacts);
        }

        private bool IsManagedCloseStartupGuardEligible(ManagedCloseStartupFacts facts)
        {
            return ManagedCloseStartupGuardPolicy.IsEligible(ToManagedCloseStartupGuardFacts(facts));
        }

        private static ManagedCloseStartupGuardFacts ToManagedCloseStartupGuardFacts(ManagedCloseStartupFacts facts)
        {
            if (facts == null)
            {
                return null;
            }

            return new ManagedCloseStartupGuardFacts
            {
                ReadFailed = facts.ReadFailed,
                WorkbookOpenObserved = facts.WorkbookOpenObserved,
                ActiveWorkbookPresent = facts.ActiveWorkbookPresent,
                WorkbooksCount = facts.WorkbooksCount,
                VisibleNonKernelWorkbookExists = facts.VisibleNonKernelWorkbookExists,
                HasOpenKernelWorkbook = facts.HasOpenKernelWorkbook,
                ApplicationVisible = facts.ApplicationVisible,
                CommandLineHasRestoreSwitch = facts.CommandLineHasRestoreSwitch,
                CommandLineHasEmbeddingSwitch = facts.CommandLineHasEmbeddingSwitch
            };
        }

        private ManagedCloseStartupFacts CaptureManagedCloseStartupFacts(string phase)
        {
            int currentProcessId = SafeGetCurrentProcessId();
            var facts = new ManagedCloseStartupFacts
            {
                Phase = phase ?? string.Empty,
                WorkbookOpenObserved = _workbookOpenObservedSinceStartup,
                ProcessId = currentProcessId,
                ProcessStartTime = SafeGetProcessStartTime(),
                CommandLine = SafeGetCommandLine()
            };

            try
            {
                facts.ApplicationVisible = _application != null && _application.Visible;
            }
            catch
            {
                facts.ReadFailed = true;
                facts.ApplicationVisibleReadFailed = true;
            }

            try
            {
                Excel.Workbook activeWorkbook = _getActiveWorkbook == null ? null : _getActiveWorkbook();
                facts.ActiveWorkbookPresent = activeWorkbook != null;
                facts.ActiveWorkbookName = _getWorkbookName == null ? string.Empty : _getWorkbookName(activeWorkbook);
            }
            catch
            {
                facts.ReadFailed = true;
                facts.ActiveWorkbookReadFailed = true;
            }

            try
            {
                facts.WorkbooksCount = _application == null || _application.Workbooks == null ? -1 : _application.Workbooks.Count;
            }
            catch
            {
                facts.ReadFailed = true;
                facts.WorkbooksCount = -1;
                facts.WorkbooksCountReadFailed = true;
            }

            try
            {
                CaptureOpenWorkbookFacts(facts);
            }
            catch
            {
                facts.ReadFailed = true;
                facts.OpenWorkbookScanFailed = true;
            }

            facts.CommandLineHasEmbeddingSwitch = ContainsCommandLineSwitch(facts.CommandLine, "embedding")
                || ContainsCommandLineSwitch(facts.CommandLine, "automation");
            facts.CommandLineHasRestoreSwitch = ContainsCommandLineSwitch(facts.CommandLine, "restore");
            return facts;
        }

        private void CaptureOpenWorkbookFacts(ManagedCloseStartupFacts facts)
        {
            if (_application == null || _application.Workbooks == null)
            {
                return;
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (workbook == null)
                {
                    continue;
                }

                bool isKernel = false;
                try
                {
                    isKernel = _isKernelWorkbook != null && _isKernelWorkbook(workbook);
                }
                catch
                {
                    facts.ReadFailed = true;
                    facts.OpenWorkbookScanFailed = true;
                }

                if (isKernel)
                {
                    facts.HasOpenKernelWorkbook = true;
                    continue;
                }

                if (WorkbookHasVisibleWindow(workbook, facts))
                {
                    facts.VisibleNonKernelWorkbookExists = true;
                }
            }
        }

        private static bool WorkbookHasVisibleWindow(Excel.Workbook workbook, ManagedCloseStartupFacts facts)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null && window.Visible)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                if (facts != null)
                {
                    facts.ReadFailed = true;
                    facts.OpenWorkbookScanFailed = true;
                }

                return false;
            }

            return false;
        }

        private void QuitEmptyStartupExcelForManagedClose(ManagedWorkbookCloseMarkerReadResult markerResult, ManagedCloseStartupFacts facts)
        {
            _logger?.Info(
                KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-guard-quit-attempt"
                + FormatManagedCloseMarkerFields(markerResult)
                + facts.ToTraceFields());
            bool previousDisplayAlerts = true;
            bool hasDisplayAlertsSnapshot = false;
            try
            {
                previousDisplayAlerts = _application.DisplayAlerts;
                hasDisplayAlertsSnapshot = true;
                _application.DisplayAlerts = false;
                _application.Quit();
                _logger?.Info(
                    KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-guard-quit-completed"
                    + FormatManagedCloseMarkerFields(markerResult)
                    + facts.ToTraceFields());
            }
            catch (Exception ex)
            {
                if (hasDisplayAlertsSnapshot)
                {
                    try
                    {
                        _application.DisplayAlerts = previousDisplayAlerts;
                    }
                    catch
                    {
                    }
                }

                _logger?.Error(
                    "Managed close startup guard quit failed."
                    + FormatManagedCloseMarkerFields(markerResult)
                    + facts.ToTraceFields(),
                    ex);
            }
        }

        private void LogManagedCloseStartupMarker(ManagedWorkbookCloseMarkerReadResult markerResult)
        {
            _logger?.Info(
                KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-marker"
                + FormatManagedCloseMarkerFields(markerResult));
        }

        private void LogManagedCloseStartupFacts(ManagedCloseStartupFacts facts, ManagedWorkbookCloseMarkerReadResult markerResult)
        {
            _logger?.Info(
                KernelFlickerTracePrefix + " source=ThisAddIn action=managed-close-startup-facts"
                + FormatManagedCloseMarkerFields(markerResult)
                + (facts == null ? string.Empty : facts.ToTraceFields()));
        }

        private static string FormatManagedCloseMarkerFields(ManagedWorkbookCloseMarkerReadResult markerResult)
        {
            if (markerResult == null)
            {
                return ", markerPresent=False, markerStatus=notConfigured";
            }

            ManagedWorkbookCloseMarker marker = markerResult.Marker;
            return ", markerPresent=" + (markerResult.Status != ManagedWorkbookCloseMarkerReadStatus.NoMarker).ToString()
                + ", markerStatus=" + markerResult.Status.ToString()
                + ", markerKind=" + (marker == null ? string.Empty : marker.Kind.ToString())
                + ", markerCreatedUtc=" + (marker == null ? string.Empty : marker.CreatedUtc.ToString("O", CultureInfo.InvariantCulture))
                + ", markerTtlSeconds=" + (marker == null ? string.Empty : marker.TimeToLiveSeconds.ToString(CultureInfo.InvariantCulture))
                + ", markerAgeMs=" + (markerResult.Age.HasValue ? ((long)markerResult.Age.Value.TotalMilliseconds).ToString(CultureInfo.InvariantCulture) : string.Empty)
                + ", markerPath=" + markerResult.MarkerPath;
        }

        private static int SafeGetCurrentProcessId()
        {
            try
            {
                return Process.GetCurrentProcess().Id;
            }
            catch
            {
                return 0;
            }
        }

        private static DateTime SafeGetProcessStartTime()
        {
            try
            {
                return Process.GetCurrentProcess().StartTime;
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        private static string SafeGetCommandLine()
        {
            try
            {
                return (Environment.CommandLine ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool ContainsCommandLineSwitch(string commandLine, string switchName)
        {
            return !string.IsNullOrWhiteSpace(commandLine)
                && !string.IsNullOrWhiteSpace(switchName)
                && commandLine.IndexOf(switchName, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal sealed class ManagedCloseStartupFacts
        {
            internal string Phase { get; set; }

            internal int ProcessId { get; set; }

            internal DateTime ProcessStartTime { get; set; }

            internal string CommandLine { get; set; }

            internal bool CommandLineHasEmbeddingSwitch { get; set; }

            internal bool CommandLineHasRestoreSwitch { get; set; }

            internal bool ApplicationVisible { get; set; }

            internal bool ActiveWorkbookPresent { get; set; }

            internal string ActiveWorkbookName { get; set; }

            internal int WorkbooksCount { get; set; }

            internal bool VisibleNonKernelWorkbookExists { get; set; }

            internal bool HasOpenKernelWorkbook { get; set; }

            internal bool WorkbookOpenObserved { get; set; }

            internal bool ReadFailed { get; set; }

            internal bool ApplicationVisibleReadFailed { get; set; }

            internal bool ActiveWorkbookReadFailed { get; set; }

            internal bool WorkbooksCountReadFailed { get; set; }

            internal bool OpenWorkbookScanFailed { get; set; }

            internal string ToTraceFields()
            {
                return ", phase=" + (Phase ?? string.Empty)
                    + ", pid=" + ProcessId.ToString(CultureInfo.InvariantCulture)
                    + ", processStartTime=" + (ProcessStartTime == DateTime.MinValue ? string.Empty : ProcessStartTime.ToString("O", CultureInfo.InvariantCulture))
                    + ", activeWorkbookPresent=" + ActiveWorkbookPresent.ToString()
                    + ", activeWorkbookName=" + (ActiveWorkbookName ?? string.Empty)
                    + ", workbooksCount=" + WorkbooksCount.ToString(CultureInfo.InvariantCulture)
                    + ", visibleNonKernelWorkbookExists=" + VisibleNonKernelWorkbookExists.ToString()
                    + ", hasOpenKernelWorkbook=" + HasOpenKernelWorkbook.ToString()
                    + ", workbookOpenObserved=" + WorkbookOpenObserved.ToString()
                    + ", applicationVisible=" + ApplicationVisible.ToString()
                    + ", commandLineHasEmbeddingSwitch=" + CommandLineHasEmbeddingSwitch.ToString()
                    + ", commandLineHasRestoreSwitch=" + CommandLineHasRestoreSwitch.ToString()
                    + ", commandLine=\"" + (CommandLine ?? string.Empty).Replace("\"", "'") + "\""
                    + ", readFailed=" + ReadFailed.ToString()
                    + ", applicationVisibleReadFailed=" + ApplicationVisibleReadFailed.ToString()
                    + ", activeWorkbookReadFailed=" + ActiveWorkbookReadFailed.ToString()
                    + ", workbooksCountReadFailed=" + WorkbooksCountReadFailed.ToString()
                    + ", openWorkbookScanFailed=" + OpenWorkbookScanFailed.ToString();
            }
        }
    }
}
