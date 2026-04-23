using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;


namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseWorkbookOpenStrategy
    {
        private const string SharedHiddenExcelEnvironmentVariableName = "CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL";
        private const string LegacyHiddenRouteName = "legacy-isolated";
        private const string SharedHiddenRouteName = "experimental-shared";
        private const string CreatedCaseDisplayHiddenRouteName = "created-case-display";
        private readonly Excel.Application _application;
        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly Logger _logger;

        internal CaseWorkbookOpenStrategy(Excel.Application application, WorkbookRoleResolver workbookRoleResolver, Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException(nameof(workbookRoleResolver));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal void RegisterKnownCasePath(string caseWorkbookPath)
        {
            _workbookRoleResolver.RegisterKnownCasePath(caseWorkbookPath);
        }

        internal void ShutdownLegacyHiddenApplication()
        {
            // legacy-isolated no longer owns a dedicated hidden Excel instance.
        }

        internal Excel.Workbook OpenVisibleWorkbook(string caseWorkbookPath)
        {
            _logger.Info("Case workbook open visible requested. path=" + (caseWorkbookPath ?? string.Empty));
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Window previousActiveWindow = null;
            bool previousScreenUpdating = _application.ScreenUpdating;
            try
            {
                previousActiveWindow = _application.ActiveWindow;
                _application.ScreenUpdating = false;
                try
                {
                    Excel.Workbook workbook = _application.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                    _logger.Info("Case workbook visible open completed. path=" + (caseWorkbookPath ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());
                    LogOpenVisibleWorkbookWindowState("after-open", workbook);
                    _workbookRoleResolver.RegisterKnownCaseWorkbook(workbook);
                    LogOpenVisibleWorkbookWindowState("before-hide", workbook);
                    HideOpenedWorkbookWindow(workbook);
                    LogOpenVisibleWorkbookWindowState("after-hide", workbook);
                    RestorePreviousWindow(previousActiveWindow);
                    LogOpenVisibleWorkbookWindowState("after-restore-previous-window", workbook);
                    _logger.Info("Case workbook visible open post-processing completed. path=" + (caseWorkbookPath ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());
                    return workbook;
                }
                finally
                {
                    try
                    {
                        _application.ScreenUpdating = previousScreenUpdating;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("OpenVisibleWorkbook failed to restore ScreenUpdating.", ex);
                    }
                }
            }
            catch
            {
                RestorePreviousWindow(previousActiveWindow);
                throw;
            }
        }

        internal HiddenCaseWorkbookSession OpenHiddenWorkbook(string caseWorkbookPath)
        {
            _logger.Info("Case workbook open hidden requested. path=" + (caseWorkbookPath ?? string.Empty));
            if (!IsSharedHiddenExcelEnabled())
            {
                _logger.Info("Case workbook hidden route selected. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + LegacyHiddenRouteName);
                return OpenHiddenWorkbookWithDedicatedApplication(caseWorkbookPath);
            }

            _logger.Info("Case workbook hidden route selected. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName);
            return OpenHiddenWorkbookWithSharedApplication(caseWorkbookPath);
        }

        internal Excel.Workbook OpenHiddenForCaseDisplay(string caseWorkbookPath)
        {
            _logger.Info("Case workbook hidden-for-display requested. path=" + (caseWorkbookPath ?? string.Empty));
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Window previousActiveWindow = null;
            Excel.Workbook workbook = null;
            bool previousScreenUpdating = _application.ScreenUpdating;
            bool previousEnableEvents = _application.EnableEvents;
            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                previousActiveWindow = _application.ActiveWindow;
                _logger.Info(
                    "Case workbook hidden-for-display Excel state captured. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + CreatedCaseDisplayHiddenRouteName
                    + ", screenUpdating="
                    + previousScreenUpdating.ToString()
                    + ", enableEvents="
                    + previousEnableEvents.ToString()
                    + ", displayAlerts="
                    + previousDisplayAlerts.ToString()
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                _application.DisplayAlerts = false;
                _logger.Info(
                    "Case workbook hidden-for-display Excel state applied. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + CreatedCaseDisplayHiddenRouteName
                    + ", screenUpdating=false, enableEvents=false, displayAlerts=false, elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                workbook = _application.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                _workbookRoleResolver.RegisterKnownCaseWorkbook(workbook);
                HideOpenedWorkbookWindow(workbook);
                RestorePreviousWindow(previousActiveWindow);
                _logger.Info(
                    "Case workbook hidden-for-display open completed. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + CreatedCaseDisplayHiddenRouteName
                    + ", appHwnd="
                    + SafeApplicationHwnd(_application)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return workbook;
            }
            catch
            {
                TryCloseWorkbookWithoutSaving(workbook);
                RestorePreviousWindow(previousActiveWindow);
                throw;
            }
            finally
            {
                RestoreSharedApplicationState(
                    caseWorkbookPath,
                    CreatedCaseDisplayHiddenRouteName,
                    stopwatch,
                    previousScreenUpdating,
                    previousEnableEvents,
                    previousDisplayAlerts);
            }
        }

        private HiddenCaseWorkbookSession OpenHiddenWorkbookWithDedicatedApplication(string caseWorkbookPath)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Window previousActiveWindow = null;
            Excel.Workbook workbook = null;
            bool previousScreenUpdating = _application.ScreenUpdating;
            bool previousEnableEvents = _application.EnableEvents;
            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                previousActiveWindow = _application.ActiveWindow;
                _logger.Info(
                    "Case workbook legacy hidden Excel state captured. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + LegacyHiddenRouteName
                    + ", screenUpdating="
                    + previousScreenUpdating.ToString()
                    + ", enableEvents="
                    + previousEnableEvents.ToString()
                    + ", displayAlerts="
                    + previousDisplayAlerts.ToString()
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                _application.DisplayAlerts = false;
                _logger.Info(
                    "Case workbook legacy hidden Excel state applied. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + LegacyHiddenRouteName
                    + ", screenUpdating=false, enableEvents=false, displayAlerts=false, elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                workbook = _application.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                HideOpenedWorkbookWindow(workbook);
                RestorePreviousWindow(previousActiveWindow);
                _logger.Info(
                    "Case workbook hidden Excel session opened. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + LegacyHiddenRouteName
                    + ", appHwnd="
                    + SafeApplicationHwnd(_application)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return new HiddenCaseWorkbookSession(
                    _application,
                    workbook,
                    LegacyHiddenRouteName,
                    closeAction: () =>
                    {
                        Stopwatch closeStopwatch = Stopwatch.StartNew();
                        _logger.Info("Case workbook hidden session close entered. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + LegacyHiddenRouteName);
                        try
                        {
                            workbook.Close(false, Type.Missing, Type.Missing);
                            _logger.Info("Case workbook hidden session workbook close completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + LegacyHiddenRouteName + ", elapsedMs=" + closeStopwatch.ElapsedMilliseconds.ToString());
                        }
                        finally
                        {
                            RestoreSharedApplicationState(caseWorkbookPath, LegacyHiddenRouteName, stopwatch, previousScreenUpdating, previousEnableEvents, previousDisplayAlerts);
                            RestorePreviousWindow(previousActiveWindow);
                        }
                    },
                    abortAction: () =>
                    {
                        try
                        {
                            workbook.Close(false, Type.Missing, Type.Missing);
                        }
                        finally
                        {
                            RestoreSharedApplicationState(caseWorkbookPath, LegacyHiddenRouteName, stopwatch, previousScreenUpdating, previousEnableEvents, previousDisplayAlerts);
                            RestorePreviousWindow(previousActiveWindow);
                        }
                    });
            }
            catch
            {
                TryCloseWorkbookWithoutSaving(workbook);
                RestoreSharedApplicationState(caseWorkbookPath, LegacyHiddenRouteName, stopwatch, previousScreenUpdating, previousEnableEvents, previousDisplayAlerts);
                RestorePreviousWindow(previousActiveWindow);
                throw;
            }
        }

        private HiddenCaseWorkbookSession OpenHiddenWorkbookWithSharedApplication(string caseWorkbookPath)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Window previousActiveWindow = null;
            Excel.Workbook workbook = null;
            bool previousScreenUpdating = _application.ScreenUpdating;
            bool previousEnableEvents = _application.EnableEvents;
            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                previousActiveWindow = _application.ActiveWindow;
                _logger.Info(
                    "Case workbook shared hidden Excel state captured. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", screenUpdating="
                    + previousScreenUpdating.ToString()
                    + ", enableEvents="
                    + previousEnableEvents.ToString()
                    + ", displayAlerts="
                    + previousDisplayAlerts.ToString()
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                _application.DisplayAlerts = false;
                _logger.Info(
                    "Case workbook shared hidden Excel state applied. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", screenUpdating=false, enableEvents=false, displayAlerts=false, elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                workbook = _application.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                HideOpenedWorkbookWindow(workbook);
                RestorePreviousWindow(previousActiveWindow);
                _logger.Info(
                    "Case workbook hidden Excel session opened. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + SharedHiddenRouteName
                    + ", appHwnd="
                    + SafeApplicationHwnd(_application)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return new HiddenCaseWorkbookSession(
                    _application,
                    workbook,
                    SharedHiddenRouteName,
                    closeAction: () =>
                    {
                        Stopwatch closeStopwatch = Stopwatch.StartNew();
                        _logger.Info("Case workbook hidden session close entered. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName);
                        try
                        {
                            _logger.Info("Case workbook hidden session inner save starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName + ", elapsedMs=" + closeStopwatch.ElapsedMilliseconds.ToString());
                            workbook.Save();
                            _logger.Info("Case workbook hidden session inner save completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName + ", elapsedMs=" + closeStopwatch.ElapsedMilliseconds.ToString());
                            _logger.Info("Case workbook hidden session workbook close starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName + ", elapsedMs=" + closeStopwatch.ElapsedMilliseconds.ToString());
                            workbook.Close(false, Type.Missing, Type.Missing);
                            _logger.Info("Case workbook hidden session workbook close completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName + ", elapsedMs=" + closeStopwatch.ElapsedMilliseconds.ToString());
                        }
                        finally
                        {
                            RestoreSharedApplicationState(caseWorkbookPath, SharedHiddenRouteName, stopwatch, previousScreenUpdating, previousEnableEvents, previousDisplayAlerts);
                            RestorePreviousWindow(previousActiveWindow);
                            _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName + ", elapsedMs=" + closeStopwatch.ElapsedMilliseconds.ToString());
                        }
                    },
                    abortAction: () =>
                    {
                        try
                        {
                            workbook.Close(false, Type.Missing, Type.Missing);
                        }
                        finally
                        {
                            RestoreSharedApplicationState(caseWorkbookPath, SharedHiddenRouteName, stopwatch, previousScreenUpdating, previousEnableEvents, previousDisplayAlerts);
                            RestorePreviousWindow(previousActiveWindow);
                        }
                    });
            }
            catch
            {
                TryCloseWorkbookWithoutSaving(workbook);
                RestoreSharedApplicationState(caseWorkbookPath, SharedHiddenRouteName, stopwatch, previousScreenUpdating, previousEnableEvents, previousDisplayAlerts);
                RestorePreviousWindow(previousActiveWindow);
                throw;
            }
        }

        private static bool IsSharedHiddenExcelEnabled()
        {
            string value = Environment.GetEnvironmentVariable(SharedHiddenExcelEnvironmentVariableName);
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private void RestoreSharedApplicationState(string caseWorkbookPath, string routeName, Stopwatch stopwatch, bool screenUpdating, bool enableEvents, bool displayAlerts)
        {
            try
            {
                _application.ScreenUpdating = screenUpdating;
                _application.EnableEvents = enableEvents;
                _application.DisplayAlerts = displayAlerts;
                _logger.Info(
                    "Case workbook hidden Excel state restored. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + (routeName ?? string.Empty)
                    + ", screenUpdating="
                    + screenUpdating.ToString()
                    + ", enableEvents="
                    + enableEvents.ToString()
                    + ", displayAlerts="
                    + displayAlerts.ToString()
                    + ", elapsedMs="
                    + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
            }
            catch (Exception ex)
            {
                _logger.Error("RestoreSharedApplicationState failed.", ex);
            }
        }

        private void TryCloseWorkbookWithoutSaving(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                workbook.Close(false, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                _logger.Error("TryCloseWorkbookWithoutSaving failed.", ex);
            }
        }

        private static string SafeApplicationHwnd(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : Convert.ToString(application.Hwnd) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void LogOpenVisibleWorkbookWindowState(string stage, Excel.Workbook openedWorkbook)
        {
            _logger.Info(
                "Case workbook open visible state. stage="
                + (stage ?? string.Empty)
                + ", appHwnd="
                + SafeApplicationHwnd(_application)
                + ", workbooksCount="
                + SafeWorkbooksCount(_application)
                + ", activeWorkbookName="
                + SafeWorkbookName(_application == null ? null : _application.ActiveWorkbook)
                + ", activeWindowCaption="
                + SafeWindowCaption(_application == null ? null : _application.ActiveWindow)
                + ", openedWorkbookWindows="
                + DescribeWorkbookWindows(openedWorkbook));
        }

        private static string SafeWorkbooksCount(Excel.Application application)
        {
            try
            {
                return application == null || application.Workbooks == null
                    ? string.Empty
                    : application.Workbooks.Count.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWorkbookName(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null ? string.Empty : workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
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
                return Convert.ToString(lateBoundWindow.Caption) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowVisible(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.Visible.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string DescribeWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return "count=0";
            }

            try
            {
                int count = workbook.Windows == null ? 0 : workbook.Windows.Count;
                string result = "count=" + count.ToString();
                for (int index = 1; index <= count; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        result += ";index="
                            + index.ToString()
                            + ",visible="
                            + SafeWindowVisible(window)
                            + ",caption="
                            + SafeWindowCaption(window)
                            + ",hwnd="
                            + SafeWindowHwnd(window);
                    }
                    catch
                    {
                        result += ";index=" + index.ToString() + ",error=window-state-unavailable";
                    }
                }

                return result;
            }
            catch
            {
                return "count=";
            }
        }

        private void HideOpenedWorkbookWindow(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null)
                    {
                        window.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("HideOpenedWorkbookWindow failed.", ex);
            }
        }

        private void RestorePreviousWindow(Excel.Window previousActiveWindow)
        {
            if (previousActiveWindow == null)
            {
                return;
            }

            try
            {
                previousActiveWindow.Visible = true;
                previousActiveWindow.Activate();
            }
            catch (Exception ex)
            {
                _logger.Error("RestorePreviousWindow failed.", ex);
            }
        }

        internal sealed class HiddenCaseWorkbookSession
        {
            private readonly Action _closeAction;
            private readonly Action _abortAction;
            private bool _closed;

            internal HiddenCaseWorkbookSession(Excel.Application application, Excel.Workbook workbook, string routeName, Action closeAction, Action abortAction)
            {
                Application = application ?? throw new ArgumentNullException(nameof(application));
                Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
                RouteName = routeName ?? string.Empty;
                _closeAction = closeAction ?? throw new ArgumentNullException(nameof(closeAction));
                _abortAction = abortAction ?? throw new ArgumentNullException(nameof(abortAction));
            }

            internal Excel.Application Application { get; }

            internal Excel.Workbook Workbook { get; }

            internal string RouteName { get; }

            internal void Close()
            {
                Execute(_closeAction);
            }

            internal void Abort()
            {
                Execute(_abortAction);
            }

            private void Execute(Action action)
            {
                if (_closed)
                {
                    return;
                }

                action();
                _closed = true;
            }
        }
    }
}
