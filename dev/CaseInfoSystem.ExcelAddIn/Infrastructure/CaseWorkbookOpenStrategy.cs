using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseWorkbookOpenStrategy
    {
        private const string SharedHiddenExcelEnvironmentVariableName = "CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL";
        private const string HiddenApplicationCacheEnvironmentVariableName = "CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE";
        private const string HiddenApplicationCacheIdleSecondsEnvironmentVariableName = "CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE_IDLE_SECONDS";
        private const string LegacyHiddenRouteName = "legacy-isolated";
        private const string SharedHiddenRouteName = "experimental-shared";
        private const string CreatedCaseDisplayHiddenRouteName = "created-case-display";
        private const string HiddenApplicationCacheRouteName = "app-cache";
        private const string HiddenApplicationCacheBypassInUseRouteName = "app-cache-bypass-inuse";
        private const int DefaultHiddenApplicationCacheIdleSeconds = 15;
        private readonly Excel.Application _application;
        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly Logger _logger;
        private readonly object _hiddenApplicationCacheSync = new object();
        private CachedHiddenApplicationSlot _cachedHiddenApplication;
        private Timer _hiddenApplicationIdleTimer;

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
            CachedHiddenApplicationSlot slotToDispose = null;
            lock (_hiddenApplicationCacheSync)
            {
                DisposeHiddenApplicationIdleTimerUnlocked();
                slotToDispose = _cachedHiddenApplication;
                _cachedHiddenApplication = null;
            }

            DisposeCachedHiddenApplicationSlot(slotToDispose, "shutdown-cleanup");
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
            if (IsHiddenApplicationCacheEnabled())
            {
                return OpenHiddenWorkbookWithApplicationCache(caseWorkbookPath);
            }

            if (!IsSharedHiddenExcelEnabled())
            {
                _logger.Info("Case workbook hidden route selected. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + LegacyHiddenRouteName);
                return OpenHiddenWorkbookWithDedicatedApplication(caseWorkbookPath);
            }

            _logger.Info("Case workbook hidden route selected. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + SharedHiddenRouteName);
            return OpenHiddenWorkbookWithSharedApplication(caseWorkbookPath);
        }

        private HiddenCaseWorkbookSession OpenHiddenWorkbookWithApplicationCache(string caseWorkbookPath)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Application hiddenApplication = null;
            Excel.Workbook workbook = null;
            bool reusedApplication = false;
            bool bypassBecauseInUse = false;
            CachedHiddenApplicationSlot expiredSlotToDispose = null;

            lock (_hiddenApplicationCacheSync)
            {
                EnsureHiddenApplicationIdleTimerUnlocked();
                expiredSlotToDispose = CleanupExpiredCachedHiddenApplicationUnlocked("OpenHiddenWorkbook.Acquire");

                if (_cachedHiddenApplication != null)
                {
                    if (_cachedHiddenApplication.IsInUse)
                    {
                        bypassBecauseInUse = true;
                    }
                    else if (!IsCachedHiddenApplicationHealthyUnlocked(_cachedHiddenApplication))
                    {
                        _logger.Warn(
                            "hidden-app-cache unhealthy. reason=acquire-health-check-failed, appHwnd="
                            + SafeApplicationHwnd(_cachedHiddenApplication.Application));
                        DisposeCachedHiddenApplicationSlotUnlocked("acquire-unhealthy");
                    }
                    else
                    {
                        hiddenApplication = _cachedHiddenApplication.Application;
                        _cachedHiddenApplication.IsInUse = true;
                        _cachedHiddenApplication.IdleSinceUtc = DateTime.MinValue;
                        StopHiddenApplicationIdleTimerUnlocked();
                        reusedApplication = true;
                    }
                }

                if (!bypassBecauseInUse && hiddenApplication == null)
                {
                    hiddenApplication = CreateDedicatedHiddenApplication(caseWorkbookPath, HiddenApplicationCacheRouteName, stopwatch);
                    _cachedHiddenApplication = new CachedHiddenApplicationSlot(hiddenApplication)
                    {
                        IsInUse = true,
                        IsOwnedByCache = true,
                        IdleSinceUtc = DateTime.MinValue
                    };
                }
            }

            DisposeCachedHiddenApplicationSlot(expiredSlotToDispose, "OpenHiddenWorkbook.Acquire");

            if (bypassBecauseInUse)
            {
                _logger.Info(
                    "hidden-app-cache bypassed because in-use. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + HiddenApplicationCacheBypassInUseRouteName
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return OpenDedicatedHiddenWorkbookSession(caseWorkbookPath, HiddenApplicationCacheBypassInUseRouteName, saveBeforeClose: false);
            }

            _logger.Info(
                "hidden-app-cache "
                + (reusedApplication ? "reused" : "created")
                + ". path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + HiddenApplicationCacheRouteName
                + ", appHwnd="
                + SafeApplicationHwnd(hiddenApplication)
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString());

            try
            {
                PrepareHiddenApplicationForUse(hiddenApplication);
                workbook = hiddenApplication.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                HideOpenedWorkbookWindow(workbook);
                _logger.Info(
                    "Case workbook hidden Excel session opened. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + HiddenApplicationCacheRouteName
                    + ", appHwnd="
                    + SafeApplicationHwnd(hiddenApplication)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return new HiddenCaseWorkbookSession(
                    hiddenApplication,
                    workbook,
                    HiddenApplicationCacheRouteName,
                    closeAction: () =>
                    {
                        Stopwatch closeStopwatch = Stopwatch.StartNew();
                        _logger.Info(
                            "Case workbook hidden session close entered. path="
                            + (caseWorkbookPath ?? string.Empty)
                            + ", route="
                            + HiddenApplicationCacheRouteName);
                        CleanupCachedHiddenSession(
                            caseWorkbookPath,
                            HiddenApplicationCacheRouteName,
                            closeStopwatch,
                            hiddenApplication,
                            workbook,
                            markPoisoned: false);
                    },
                    abortAction: () =>
                    {
                        CleanupCachedHiddenSession(
                            caseWorkbookPath,
                            HiddenApplicationCacheRouteName,
                            stopwatch,
                            hiddenApplication,
                            workbook,
                            markPoisoned: true);
                    });
            }
            catch
            {
                TryCloseWorkbookWithoutSaving(workbook);
                ReleaseComObject(workbook);
                CleanupCachedHiddenSession(
                    caseWorkbookPath,
                    HiddenApplicationCacheRouteName,
                    stopwatch,
                    hiddenApplication,
                    workbook: null,
                    markPoisoned: true);
                throw;
            }
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
            return OpenDedicatedHiddenWorkbookSession(caseWorkbookPath, LegacyHiddenRouteName, saveBeforeClose: false);
        }

        private HiddenCaseWorkbookSession OpenHiddenWorkbookWithSharedApplication(string caseWorkbookPath)
        {
            return OpenDedicatedHiddenWorkbookSession(caseWorkbookPath, SharedHiddenRouteName, saveBeforeClose: true);
        }

        private HiddenCaseWorkbookSession OpenDedicatedHiddenWorkbookSession(string caseWorkbookPath, string routeName, bool saveBeforeClose)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Application hiddenApplication = null;
            Excel.Workbook workbook = null;
            try
            {
                hiddenApplication = CreateDedicatedHiddenApplication(caseWorkbookPath, routeName, stopwatch);
                workbook = hiddenApplication.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                HideOpenedWorkbookWindow(workbook);
                _logger.Info(
                    "Case workbook hidden Excel session opened. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + routeName
                    + ", appHwnd="
                    + SafeApplicationHwnd(hiddenApplication)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return new HiddenCaseWorkbookSession(
                    hiddenApplication,
                    workbook,
                    routeName,
                    closeAction: () =>
                    {
                        Stopwatch closeStopwatch = Stopwatch.StartNew();
                        _logger.Info("Case workbook hidden session close entered. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + routeName);
                        CleanupDedicatedHiddenSession(caseWorkbookPath, routeName, closeStopwatch, hiddenApplication, workbook, saveBeforeClose);
                    },
                    abortAction: () =>
                    {
                        CleanupDedicatedHiddenSession(caseWorkbookPath, routeName, stopwatch, hiddenApplication, workbook, saveBeforeClose: false);
                    });
            }
            catch
            {
                CleanupDedicatedHiddenSession(caseWorkbookPath, routeName, stopwatch, hiddenApplication, workbook, saveBeforeClose: false);
                throw;
            }
        }

        private static bool IsSharedHiddenExcelEnabled()
        {
            string value = Environment.GetEnvironmentVariable(SharedHiddenExcelEnvironmentVariableName);
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHiddenApplicationCacheEnabled()
        {
            string value = Environment.GetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName);
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static int ResolveHiddenApplicationCacheIdleSeconds()
        {
            string value = Environment.GetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName);
            if (int.TryParse(value, out int parsed) && parsed > 0)
            {
                return parsed;
            }

            return DefaultHiddenApplicationCacheIdleSeconds;
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

        private static void PrepareHiddenApplicationForUse(Excel.Application hiddenApplication)
        {
            if (hiddenApplication == null)
            {
                return;
            }

            hiddenApplication.Visible = false;
            hiddenApplication.DisplayAlerts = false;
            hiddenApplication.ScreenUpdating = false;
            hiddenApplication.UserControl = false;
            hiddenApplication.EnableEvents = false;
        }

        private Excel.Application CreateDedicatedHiddenApplication(string caseWorkbookPath, string routeName, Stopwatch stopwatch)
        {
            Excel.Application hiddenApplication = new Excel.Application();
            try
            {
                PrepareHiddenApplicationForUse(hiddenApplication);
                _logger.Info(
                    "Case workbook dedicated hidden Excel created. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + (routeName ?? string.Empty)
                    + ", visible=false, displayAlerts=false, screenUpdating=false, userControl=false, enableEvents=false, elapsedMs="
                    + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                return hiddenApplication;
            }
            catch
            {
                TryQuitApplication(hiddenApplication);
                ReleaseComObject(hiddenApplication);
                throw;
            }
        }

        private void CleanupDedicatedHiddenSession(string caseWorkbookPath, string routeName, Stopwatch stopwatch, Excel.Application application, Excel.Workbook workbook, bool saveBeforeClose)
        {
            try
            {
                if (saveBeforeClose && workbook != null)
                {
                    _logger.Info("Case workbook hidden session inner save starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    workbook.Save();
                    _logger.Info("Case workbook hidden session inner save completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                }

                if (workbook != null)
                {
                    _logger.Info("Case workbook hidden session workbook close starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    workbook.Close(false, Type.Missing, Type.Missing);
                    _logger.Info("Case workbook hidden session workbook close completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                }
            }
            finally
            {
                TryQuitApplication(application);
                ReleaseComObject(workbook);
                ReleaseComObject(application);
                _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
            }
        }

        private void CleanupCachedHiddenSession(string caseWorkbookPath, string routeName, Stopwatch stopwatch, Excel.Application application, Excel.Workbook workbook, bool markPoisoned)
        {
            bool closeFailed = false;
            try
            {
                if (workbook != null)
                {
                    if (markPoisoned)
                    {
                        TryCloseWorkbookWithoutSaving(workbook);
                    }
                    else
                    {
                        _logger.Info("Case workbook hidden session workbook close starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                        workbook.Close(false, Type.Missing, Type.Missing);
                        _logger.Info("Case workbook hidden session workbook close completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    }
                }
            }
            catch (Exception ex)
            {
                closeFailed = true;
                markPoisoned = true;
                _logger.Error("CleanupCachedHiddenSession workbook close failed.", ex);
            }
            finally
            {
                ReleaseComObject(workbook);
            }

            if (markPoisoned || closeFailed)
            {
                MarkCachedHiddenApplicationPoisoned(application, caseWorkbookPath, routeName, stopwatch);
                _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", cached=False, elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                return;
            }

            if (TryReturnCachedHiddenApplicationToIdle(application, caseWorkbookPath, routeName, stopwatch))
            {
                _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", cached=True, elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                return;
            }

            MarkCachedHiddenApplicationPoisoned(application, caseWorkbookPath, routeName, stopwatch);
            _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", cached=False, elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
        }

        private bool TryReturnCachedHiddenApplicationToIdle(Excel.Application application, string caseWorkbookPath, string routeName, Stopwatch stopwatch)
        {
            lock (_hiddenApplicationCacheSync)
            {
                if (_cachedHiddenApplication == null
                    || !ReferenceEquals(_cachedHiddenApplication.Application, application)
                    || !_cachedHiddenApplication.IsOwnedByCache)
                {
                    return false;
                }

                if (!IsHiddenApplicationCacheEnabled())
                {
                    _logger.Info("hidden-app-cache disabled before return-to-idle. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", appHwnd=" + SafeApplicationHwnd(application));
                    _cachedHiddenApplication.IsPoisoned = true;
                    return false;
                }

                try
                {
                    PrepareHiddenApplicationForUse(application);
                }
                catch (Exception ex)
                {
                    _logger.Error("hidden-app-cache failed to reapply hidden state.", ex);
                    _cachedHiddenApplication.IsPoisoned = true;
                    return false;
                }

                if (!IsCachedHiddenApplicationHealthyUnlocked(_cachedHiddenApplication))
                {
                    _cachedHiddenApplication.IsPoisoned = true;
                    _logger.Warn("hidden-app-cache unhealthy. reason=return-to-idle-health-check-failed, appHwnd=" + SafeApplicationHwnd(application));
                    return false;
                }

                _cachedHiddenApplication.IsInUse = false;
                _cachedHiddenApplication.IdleSinceUtc = DateTime.UtcNow;
                ScheduleHiddenApplicationIdleTimerUnlocked();
                _logger.Info("hidden-app-cache returned-to-idle. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", appHwnd=" + SafeApplicationHwnd(application) + ", idleTimeoutSeconds=" + ResolveHiddenApplicationCacheIdleSeconds().ToString() + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                return true;
            }
        }

        private void MarkCachedHiddenApplicationPoisoned(Excel.Application application, string caseWorkbookPath, string routeName, Stopwatch stopwatch)
        {
            CachedHiddenApplicationSlot slotToDispose = null;
            lock (_hiddenApplicationCacheSync)
            {
                if (_cachedHiddenApplication != null && ReferenceEquals(_cachedHiddenApplication.Application, application))
                {
                    _cachedHiddenApplication.IsPoisoned = true;
                    _logger.Warn("hidden-app-cache poisoned. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", appHwnd=" + SafeApplicationHwnd(application) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    slotToDispose = _cachedHiddenApplication;
                    _cachedHiddenApplication = null;
                    StopHiddenApplicationIdleTimerUnlocked();
                }
            }

            DisposeCachedHiddenApplicationSlot(slotToDispose, "poisoned");
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

        private void TryQuitApplication(Excel.Application application)
        {
            if (application == null)
            {
                return;
            }

            try
            {
                application.Quit();
            }
            catch (Exception ex)
            {
                _logger.Error("TryQuitApplication failed.", ex);
            }
        }

        private void HiddenApplicationIdleTimer_Tick(object sender, EventArgs e)
        {
            if (!IsHiddenApplicationCacheEnabled())
            {
                CachedHiddenApplicationSlot slotToDispose = null;
                lock (_hiddenApplicationCacheSync)
                {
                    slotToDispose = _cachedHiddenApplication;
                    _cachedHiddenApplication = null;
                    DisposeHiddenApplicationIdleTimerUnlocked();
                }

                DisposeCachedHiddenApplicationSlot(slotToDispose, "feature-flag-disabled");
                return;
            }

            CleanupExpiredCachedHiddenApplication("idle-timeout");
        }

        private void CleanupExpiredCachedHiddenApplication(string reason)
        {
            CachedHiddenApplicationSlot slotToDispose = null;
            lock (_hiddenApplicationCacheSync)
            {
                slotToDispose = CleanupExpiredCachedHiddenApplicationUnlocked(reason);
            }

            DisposeCachedHiddenApplicationSlot(slotToDispose, reason);
        }

        private CachedHiddenApplicationSlot CleanupExpiredCachedHiddenApplicationUnlocked(string reason)
        {
            if (_cachedHiddenApplication == null)
            {
                StopHiddenApplicationIdleTimerUnlocked();
                return null;
            }

            if (_cachedHiddenApplication.IsInUse)
            {
                StopHiddenApplicationIdleTimerUnlocked();
                return null;
            }

            if (_cachedHiddenApplication.IsPoisoned)
            {
                CachedHiddenApplicationSlot poisonedSlot = _cachedHiddenApplication;
                _cachedHiddenApplication = null;
                StopHiddenApplicationIdleTimerUnlocked();
                return poisonedSlot;
            }

            DateTime idleSinceUtc = _cachedHiddenApplication.IdleSinceUtc;
            if (idleSinceUtc == DateTime.MinValue)
            {
                idleSinceUtc = DateTime.UtcNow;
                _cachedHiddenApplication.IdleSinceUtc = idleSinceUtc;
            }

            if ((DateTime.UtcNow - idleSinceUtc).TotalSeconds < ResolveHiddenApplicationCacheIdleSeconds())
            {
                return null;
            }

            CachedHiddenApplicationSlot expiredSlot = _cachedHiddenApplication;
            _cachedHiddenApplication = null;
            StopHiddenApplicationIdleTimerUnlocked();
            _logger.Info("hidden-app-cache timed-out. reason=" + (reason ?? string.Empty) + ", appHwnd=" + SafeApplicationHwnd(expiredSlot.Application));
            return expiredSlot;
        }

        private void DisposeCachedHiddenApplicationSlot(CachedHiddenApplicationSlot slot, string reason)
        {
            if (slot == null)
            {
                return;
            }

            if (!slot.IsOwnedByCache)
            {
                _logger.Warn("hidden-app-cache cleanup skipped because slot is not cache-owned. reason=" + (reason ?? string.Empty) + ", appHwnd=" + SafeApplicationHwnd(slot.Application));
                return;
            }

            TryQuitApplication(slot.Application);
            ReleaseComObject(slot.Application);
            _logger.Info("hidden-app-cache discarded. reason=" + (reason ?? string.Empty) + ", appHwnd=" + SafeApplicationHwnd(slot.Application));
        }

        private void DisposeCachedHiddenApplicationSlotUnlocked(string reason)
        {
            CachedHiddenApplicationSlot slot = _cachedHiddenApplication;
            _cachedHiddenApplication = null;
            StopHiddenApplicationIdleTimerUnlocked();
            DisposeCachedHiddenApplicationSlot(slot, reason);
        }

        private void EnsureHiddenApplicationIdleTimerUnlocked()
        {
            if (_hiddenApplicationIdleTimer != null)
            {
                return;
            }

            _hiddenApplicationIdleTimer = new Timer();
            _hiddenApplicationIdleTimer.Interval = 1000;
            _hiddenApplicationIdleTimer.Tick += HiddenApplicationIdleTimer_Tick;
        }

        private void ScheduleHiddenApplicationIdleTimerUnlocked()
        {
            EnsureHiddenApplicationIdleTimerUnlocked();
            _hiddenApplicationIdleTimer.Stop();
            _hiddenApplicationIdleTimer.Start();
        }

        private void StopHiddenApplicationIdleTimerUnlocked()
        {
            if (_hiddenApplicationIdleTimer == null)
            {
                return;
            }

            _hiddenApplicationIdleTimer.Stop();
        }

        private void DisposeHiddenApplicationIdleTimerUnlocked()
        {
            if (_hiddenApplicationIdleTimer == null)
            {
                return;
            }

            _hiddenApplicationIdleTimer.Stop();
            _hiddenApplicationIdleTimer.Dispose();
            _hiddenApplicationIdleTimer = null;
        }

        private bool IsCachedHiddenApplicationHealthyUnlocked(CachedHiddenApplicationSlot slot)
        {
            if (slot == null || slot.Application == null || slot.IsPoisoned)
            {
                return false;
            }

            try
            {
                Excel.Application application = slot.Application;
                return application.Workbooks != null
                    && application.Workbooks.Count == 0
                    && application.Ready
                    && !application.Visible
                    && !application.DisplayAlerts
                    && !application.ScreenUpdating
                    && !application.EnableEvents
                    && !application.UserControl;
            }
            catch
            {
                return false;
            }
        }

        private void ReleaseComObject(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                _logger.Error("ReleaseComObject failed.", ex);
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

        private sealed class CachedHiddenApplicationSlot
        {
            internal CachedHiddenApplicationSlot(Excel.Application application)
            {
                Application = application;
            }

            internal Excel.Application Application { get; }

            internal bool IsInUse { get; set; }

            internal bool IsPoisoned { get; set; }

            internal bool IsOwnedByCache { get; set; }

            internal DateTime IdleSinceUtc { get; set; }
        }
    }
}
