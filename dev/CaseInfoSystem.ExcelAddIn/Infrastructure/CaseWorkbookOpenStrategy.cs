using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;
using Excel = Microsoft.Office.Interop.Excel;


namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseWorkbookOpenStrategy
    {
        private readonly Excel.Application _application;
        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly Logger _logger;
        private readonly CaseWorkbookOpenRouteDecisionService _routeDecisionService;
        private readonly CaseWorkbookOpenCleanupOutcomeService _cleanupOutcomeService;
        private readonly CaseWorkbookPresentationHandoffService _presentationHandoffService;
        private readonly CaseWorkbookHiddenAppLifecycleSupportService _hiddenAppLifecycleSupportService;
        private readonly Func<Excel.Application> _hiddenApplicationFactory;
        private readonly Action<object> _releaseComObject;
        private readonly object _hiddenApplicationCacheSync = new object();
        private CachedHiddenApplicationSlot _cachedHiddenApplication;
        private Timer _hiddenApplicationIdleTimer;

        internal CaseWorkbookOpenStrategy(
            Excel.Application application,
            WorkbookRoleResolver workbookRoleResolver,
            Logger logger,
            Func<Excel.Application> hiddenApplicationFactory = null,
            Action<object> releaseComObject = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException(nameof(workbookRoleResolver));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _routeDecisionService = new CaseWorkbookOpenRouteDecisionService();
            _cleanupOutcomeService = new CaseWorkbookOpenCleanupOutcomeService(_routeDecisionService);
            _presentationHandoffService = new CaseWorkbookPresentationHandoffService();
            _hiddenAppLifecycleSupportService = new CaseWorkbookHiddenAppLifecycleSupportService();
            _hiddenApplicationFactory = hiddenApplicationFactory ?? (() => new Excel.Application());
            _releaseComObject = releaseComObject;
        }

        internal void RegisterKnownCasePath(string caseWorkbookPath)
        {
            _workbookRoleResolver.RegisterKnownCasePath(caseWorkbookPath);
        }

        internal void ShutdownHiddenApplicationCache()
        {
            CachedHiddenApplicationSlot slotToDispose = null;
            lock (_hiddenApplicationCacheSync)
            {
                DisposeHiddenApplicationIdleTimerUnlocked();
                slotToDispose = _cachedHiddenApplication;
                _cachedHiddenApplication = null;
            }

            LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionShutdown,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.ShutdownHiddenApplicationCache",
                "shutdown-cleanup")
            {
                EventOutcome = slotToDispose == null ? "no-retained-instance" : "dispose-retained-instance",
                CacheEvent = "shutdown",
                RetainedInstancePresent = slotToDispose != null && slotToDispose.Application != null,
                AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(slotToDispose == null ? null : slotToDispose.Application),
                CleanupReason = "shutdown-cleanup",
                SafetyAction = slotToDispose == null ? "none" : "dispose-retained-application",
                ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            });
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
                    _logger.Info(_presentationHandoffService.BuildVisibleOpenWindowFactsMessage("after-open", _application, workbook));
                    _workbookRoleResolver.RegisterKnownCaseWorkbook(workbook);
                    _logger.Info(_presentationHandoffService.BuildVisibleOpenWindowFactsMessage("before-hide", _application, workbook));
                    HideOpenedWorkbookWindow(workbook);
                    _logger.Info(_presentationHandoffService.BuildVisibleOpenWindowFactsMessage("after-hide", _application, workbook));
                    RestorePreviousWindow(previousActiveWindow);
                    _logger.Info(_presentationHandoffService.BuildVisibleOpenWindowFactsMessage("after-restore-previous-window", _application, workbook));
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
            CaseWorkbookOpenRouteDecision routeDecision = _routeDecisionService.DecideHiddenCreateRoute();
            _logger.Info(
                "Case workbook hidden route selected. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", "
                + routeDecision.RouteTraceDetails);
            if (routeDecision.UseHiddenApplicationCache)
            {
                return OpenHiddenWorkbookWithApplicationCache(caseWorkbookPath);
            }

            return OpenDedicatedHiddenWorkbookSession(
                caseWorkbookPath,
                routeDecision.RouteName,
                routeDecision.SaveBeforeClose);
        }

        private HiddenCaseWorkbookSession OpenHiddenWorkbookWithApplicationCache(string caseWorkbookPath)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Application hiddenApplication = null;
            Excel.Workbook workbook = null;
            bool reusedApplication = false;
            bool bypassBecauseInUse = false;
            string acquisitionReason = "cache-empty";
            CaseWorkbookOpenRouteDecision bypassRouteDecision = null;
            CachedHiddenApplicationSlot expiredSlotToDispose = null;

            lock (_hiddenApplicationCacheSync)
            {
                EnsureHiddenApplicationIdleTimerUnlocked();
                expiredSlotToDispose = CleanupExpiredCachedHiddenApplicationUnlocked("OpenHiddenWorkbook.Acquire");

                if (_cachedHiddenApplication != null)
                {
                    if (_cachedHiddenApplication.IsInUse)
                    {
                        bypassRouteDecision = _routeDecisionService.DecideHiddenApplicationCacheAcquisition(cachedApplicationInUse: true);
                        bypassBecauseInUse = bypassRouteDecision.IsFallbackRoute;
                    }
                    else
                    {
                        CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts = CaptureCachedHiddenApplicationLifecycleFactsUnlocked(_cachedHiddenApplication);
                        if (!lifecycleFacts.IsReusable)
                        {
                            _logger.Warn(_hiddenAppLifecycleSupportService.BuildCacheUnhealthyMessage("acquire-health-check-failed", lifecycleFacts.AppHwnd));
                            LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionOrphanSuspicion,
                                caseWorkbookPath,
                                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                                "acquire-health-check-failed")
                            {
                                EventOutcome = "suspected-unhealthy-retained-instance",
                                CacheEvent = "acquire-health-check",
                                AppHwnd = lifecycleFacts.AppHwnd,
                                SafetyAction = "dispose-retained-application",
                                AbandonedOperation = "reuse-retained-application",
                                LifecycleFacts = lifecycleFacts,
                                ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
                            });
                            acquisitionReason = "cache-replaced-after-unhealthy";
                            DisposeCachedHiddenApplicationSlotUnlocked("acquire-unhealthy");
                        }
                        else
                        {
                            hiddenApplication = _cachedHiddenApplication.Application;
                            _cachedHiddenApplication.IsInUse = true;
                            _cachedHiddenApplication.IdleSinceUtc = DateTime.MinValue;
                            StopHiddenApplicationIdleTimerUnlocked();
                            reusedApplication = true;
                            acquisitionReason = "cache-reusable";
                        }
                    }
                }

                if (!bypassBecauseInUse && hiddenApplication == null)
                {
                    hiddenApplication = CreateDedicatedHiddenApplication(caseWorkbookPath, CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName, stopwatch);
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
                CaseWorkbookOpenRouteDecision fallbackDecision = bypassRouteDecision
                    ?? _routeDecisionService.DecideHiddenApplicationCacheAcquisition(cachedApplicationInUse: true);
                LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                    CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionFallback,
                    caseWorkbookPath,
                    fallbackDecision.RouteName,
                    "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                    fallbackDecision.Reason)
                {
                    EventOutcome = "fallback-to-dedicated-hidden-session",
                    CacheEvent = "acquire-fallback",
                    FallbackRoute = fallbackDecision.RouteName,
                    AbandonedOperation = "retained-cache-acquire",
                    SafetyAction = "open-dedicated-hidden-session",
                    ElapsedMilliseconds = stopwatch.ElapsedMilliseconds,
                    ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(fallbackDecision.RouteName)
                });
                _logger.Info(_hiddenAppLifecycleSupportService.BuildCacheBypassInUseMessage(caseWorkbookPath, fallbackDecision, stopwatch.ElapsedMilliseconds));
                return OpenDedicatedHiddenWorkbookSession(
                    caseWorkbookPath,
                    fallbackDecision.RouteName,
                    fallbackDecision.SaveBeforeClose);
            }

            LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionAcquire,
                caseWorkbookPath,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                acquisitionReason)
            {
                EventOutcome = "acquired",
                CacheEvent = "acquire",
                AcquisitionKind = reusedApplication ? "reused" : "created",
                ReusedApplication = reusedApplication,
                AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(hiddenApplication),
                ElapsedMilliseconds = stopwatch.ElapsedMilliseconds,
                ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            });
            _logger.Info(_hiddenAppLifecycleSupportService.BuildCacheAcquiredMessage(caseWorkbookPath, reusedApplication, CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName, _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(hiddenApplication), stopwatch.ElapsedMilliseconds));
            NewCaseVisibilityObservation.Log(
                _logger,
                null,
                hiddenApplication,
                null,
                null,
                "isolated-excel-acquired",
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                caseWorkbookPath,
                _hiddenAppLifecycleSupportService.BuildAcquiredObservationDetails(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName, reusedApplication, _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)));

            try
            {
                PrepareHiddenApplicationForUse(hiddenApplication);
                workbook = hiddenApplication.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                HideOpenedWorkbookWindow(workbook);
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    hiddenApplication,
                    workbook,
                    null,
                    "hidden-session-workbook-opened",
                    "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                    caseWorkbookPath,
                    "route=" + CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName);
                _logger.Info(
                    "Case workbook hidden Excel session opened. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName
                    + ", appHwnd="
                    + _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(hiddenApplication)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
                return new HiddenCaseWorkbookSession(
                    hiddenApplication,
                    workbook,
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                    closeAction: () =>
                    {
                        Stopwatch closeStopwatch = Stopwatch.StartNew();
                        _logger.Info(
                            "Case workbook hidden session close entered. path="
                            + (caseWorkbookPath ?? string.Empty)
                            + ", route="
                            + CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName);
                        CleanupCachedHiddenSession(
                            caseWorkbookPath,
                            CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                            closeStopwatch,
                            hiddenApplication,
                            workbook,
                            markPoisoned: false);
                    },
                    abortAction: () =>
                    {
                        CleanupCachedHiddenSession(
                            caseWorkbookPath,
                            CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                            stopwatch,
                            hiddenApplication,
                            workbook,
                            markPoisoned: true);
                    });
            }
            catch (Exception ex)
            {
                LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                    CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionPoisonMark,
                    caseWorkbookPath,
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                    "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                    "workbook-open-failed")
                {
                    EventOutcome = "poison-requested",
                    CacheEvent = "acquire-open-failed",
                    PoisonReason = "workbookOpenFailed",
                    AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(hiddenApplication),
                    ExceptionType = ex.GetType().Name,
                    SafetyAction = "poison-dispose",
                    ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
                });
                TryCloseWorkbookWithoutSaving(workbook);
                ReleaseComObject(workbook);
                CleanupCachedHiddenSession(
                    caseWorkbookPath,
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
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
            CaseWorkbookOpenRouteDecision routeDecision = _routeDecisionService.DecideCreatedCaseDisplayRoute();
            Stopwatch stopwatch = Stopwatch.StartNew();
            CaseWorkbookPresentationHandoffPlan handoffPlan = null;
            Excel.Workbook workbook = null;
            bool previousApplicationVisible = _presentationHandoffService.CaptureApplicationVisible(_application);
            bool previousScreenUpdating = _application.ScreenUpdating;
            bool previousEnableEvents = _application.EnableEvents;
            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                handoffPlan = _presentationHandoffService.CreateHiddenForDisplayPlan(
                    caseWorkbookPath,
                    routeDecision,
                    _application.ActiveWindow,
                    previousApplicationVisible,
                    previousScreenUpdating,
                    previousEnableEvents,
                    previousDisplayAlerts);
                _logger.Info(_presentationHandoffService.BuildHiddenForDisplayStateCapturedMessage(
                    handoffPlan,
                    stopwatch.ElapsedMilliseconds));
                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                _application.DisplayAlerts = false;
                _logger.Info(_presentationHandoffService.BuildHiddenForDisplayStateAppliedMessage(
                    handoffPlan,
                    stopwatch.ElapsedMilliseconds));
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    _application,
                    null,
                    null,
                    "shared-display-state-applied",
                    "CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay",
                    caseWorkbookPath,
                    _presentationHandoffService.BuildSharedDisplayStateAppliedObservationDetails(handoffPlan));
                Stopwatch stopwatch2 = Stopwatch.StartNew();
                workbook = _application.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "hiddenOpenToWindowVisible", "workbooksOpen", stopwatch2.ElapsedMilliseconds, "route=" + routeDecision.RouteName);
                _workbookRoleResolver.RegisterKnownCaseWorkbook(workbook);
                stopwatch2 = Stopwatch.StartNew();
                HideOpenedWorkbookWindow(workbook);
                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "hiddenOpenToWindowVisible", "hideOpenedWorkbookWindow", stopwatch2.ElapsedMilliseconds, "route=" + routeDecision.RouteName);
                stopwatch2 = Stopwatch.StartNew();
                RestorePreviousWindowForHiddenDisplay(handoffPlan, stopwatch);
                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "hiddenOpenToWindowVisible", "restorePreviousWindow", stopwatch2.ElapsedMilliseconds, "route=" + routeDecision.RouteName);
                _logger.Info(_presentationHandoffService.BuildHiddenForDisplayOpenCompletedMessage(
                    handoffPlan,
                    _presentationHandoffService.CaptureApplicationHwnd(_application),
                    stopwatch.ElapsedMilliseconds));
                return workbook;
            }
            catch
            {
                TryCloseWorkbookWithoutSaving(workbook);
                RestorePreviousWindowForHiddenDisplay(handoffPlan, stopwatch);
                throw;
            }
            finally
            {
                RestoreSharedApplicationState(handoffPlan, stopwatch);
            }
        }

        private void RestorePreviousWindowForHiddenDisplay(CaseWorkbookPresentationHandoffPlan handoffPlan, Stopwatch stopwatch)
        {
            if (handoffPlan == null)
            {
                return;
            }

            if (!handoffPlan.PreviousWindowRestoreDecision.ShouldRestore)
            {
                _logger.Info(_presentationHandoffService.BuildPreviousWindowRestoreSkippedMessage(
                    handoffPlan,
                    stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds));
                return;
            }

            RestorePreviousWindow(handoffPlan.SharedStateFacts.PreviousActiveWindow);
        }

        private HiddenCaseWorkbookSession OpenDedicatedHiddenWorkbookSession(string caseWorkbookPath, string routeName, bool saveBeforeClose)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Application hiddenApplication = null;
            Excel.Workbook workbook = null;
            try
            {
                hiddenApplication = CreateDedicatedHiddenApplication(caseWorkbookPath, routeName, stopwatch);
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    hiddenApplication,
                    null,
                    null,
                    "isolated-workbook-open-started",
                    "CaseWorkbookOpenStrategy.OpenDedicatedHiddenWorkbookSession",
                    caseWorkbookPath,
                    "scope=isolated,route=" + routeName
                    + ",openReason=hidden-create-session,"
                    + _routeDecisionService.BuildApplicationOwnerFacts(routeName));
                workbook = hiddenApplication.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    hiddenApplication,
                    workbook,
                    null,
                    "isolated-workbook-open-completed",
                    "CaseWorkbookOpenStrategy.OpenDedicatedHiddenWorkbookSession",
                    caseWorkbookPath,
                    "scope=isolated,route=" + routeName
                    + ",openReason=hidden-create-session,"
                    + _routeDecisionService.BuildApplicationOwnerFacts(routeName));
                HideOpenedWorkbookWindow(workbook);
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    hiddenApplication,
                    workbook,
                    null,
                    "hidden-session-workbook-opened",
                    "CaseWorkbookOpenStrategy.OpenDedicatedHiddenWorkbookSession",
                    caseWorkbookPath,
                    "scope=isolated,route=" + routeName
                    + ",windowHideReason=HideOpenedWorkbookWindow,"
                    + _routeDecisionService.BuildApplicationOwnerFacts(routeName));
                _logger.Info(
                    "Case workbook hidden Excel session opened. path="
                    + (caseWorkbookPath ?? string.Empty)
                    + ", route="
                    + routeName
                    + ", appHwnd="
                    + _presentationHandoffService.CaptureApplicationHwnd(hiddenApplication)
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

        private void RestoreSharedApplicationState(CaseWorkbookPresentationHandoffPlan handoffPlan, Stopwatch stopwatch)
        {
            if (handoffPlan == null)
            {
                return;
            }

            CaseWorkbookSharedDisplayStateFacts facts = handoffPlan.SharedStateFacts;
            try
            {
                _application.ScreenUpdating = facts.PreviousScreenUpdating;
                _application.EnableEvents = facts.PreviousEnableEvents;
                _application.DisplayAlerts = facts.PreviousDisplayAlerts;
                _logger.Info(_presentationHandoffService.BuildSharedDisplayStateRestoredMessage(
                    handoffPlan,
                    stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds));
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    _application,
                    null,
                    null,
                    "shared-display-state-restored",
                    "CaseWorkbookOpenStrategy.RestoreSharedApplicationState",
                    handoffPlan.CaseWorkbookPath,
                    _presentationHandoffService.BuildSharedDisplayStateRestoredObservationDetails(handoffPlan));
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
            Excel.Application hiddenApplication = _hiddenApplicationFactory();
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
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    hiddenApplication,
                    null,
                    null,
                    "isolated-excel-created",
                    "CaseWorkbookOpenStrategy.CreateDedicatedHiddenApplication",
                    caseWorkbookPath,
                    "scope=isolated,route=" + (routeName ?? string.Empty)
                    + "," + _routeDecisionService.BuildApplicationOwnerFacts(routeName));
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
	            Stopwatch stopwatch2 = Stopwatch.StartNew();
                bool workbookPresent = workbook != null;
                bool workbookCloseAttempted = false;
                bool workbookCloseCompleted = !workbookPresent;
                bool appPresent = application != null;
                bool appQuitAttempted = false;
                bool appQuitCompleted = !appPresent;
                bool cleanupFailed = false;
	            try
	            {
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        workbook,
                        null,
                        "isolated-cleanup-enter",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty) + ",saveBeforeClose=" + saveBeforeClose.ToString());
	                if (saveBeforeClose && workbook != null)
                {
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        workbook,
                        null,
                        "isolated-inner-save-started",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty));
                    _logger.Info("Case workbook hidden session inner save starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    workbook.Save();
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        workbook,
                        null,
                        "isolated-inner-save-completed",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty));
                    _logger.Info("Case workbook hidden session inner save completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                }

                if (workbook != null)
                {
                    workbookCloseAttempted = true;
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        workbook,
                        null,
                        "isolated-workbook-close-started",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty));
                    _logger.Info("Case workbook hidden session workbook close starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave(
                        workbook,
                        _logger,
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession route=" + (routeName ?? string.Empty));
                    workbookCloseCompleted = true;
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        null,
                        null,
                        "isolated-workbook-close-completed",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty));
                    _logger.Info("Case workbook hidden session workbook close completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                }
            }
                catch
                {
                    cleanupFailed = true;
                    throw;
                }
	            finally
	            {
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        null,
                        null,
                        "post-create-cleanup",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty) + ",phase=after-workbook-close-before-app-quit");
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        null,
                        null,
                        "isolated-app-quit-started",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty));
                    appQuitAttempted = application != null;
	                appQuitCompleted = TryQuitApplication(application);
                    NewCaseVisibilityObservation.Log(
                        _logger,
                        null,
                        application,
                        null,
                        null,
                        "isolated-app-quit-completed",
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        "scope=isolated,route=" + (routeName ?? string.Empty) + ",quitSucceeded=" + appQuitCompleted.ToString());
	                Stopwatch stopwatch3 = Stopwatch.StartNew();
	                ReleaseComObject(workbook);
	                ReleaseComObject(application);
                    LogHiddenCleanupOutcome(_cleanupOutcomeService.CreateDedicatedHiddenSessionOutcome(
                        "CaseWorkbookOpenStrategy.CleanupDedicatedHiddenSession",
                        caseWorkbookPath,
                        routeName,
                        new CaseWorkbookOpenCleanupFacts(
                            workbookPresent,
                            workbookCloseAttempted,
                            workbookCloseCompleted,
                            appPresent,
                            appQuitAttempted,
                            appQuitCompleted),
                        cleanupFailed));
	                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "waitUiShownToCaseCreated", "comRelease", stopwatch3.ElapsedMilliseconds, "route=" + (routeName ?? string.Empty));
	                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "waitUiShownToCaseCreated", "hiddenSessionClose", stopwatch2.ElapsedMilliseconds, "route=" + (routeName ?? string.Empty));
	                _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
	            }
	        }

	        private void CleanupCachedHiddenSession(string caseWorkbookPath, string routeName, Stopwatch stopwatch, Excel.Application application, Excel.Workbook workbook, bool markPoisoned)
	        {
	            Stopwatch stopwatch2 = Stopwatch.StartNew();
	            bool closeFailed = false;
                string closeFailureType = string.Empty;
                bool workbookPresent = workbook != null;
                bool workbookCloseAttempted = false;
                bool workbookCloseCompleted = !workbookPresent;
	            try
	            {
                if (workbook != null)
                {
                    workbookCloseAttempted = true;
                    if (markPoisoned)
                    {
                        workbookCloseCompleted = TryCloseWorkbookWithoutSaving(workbook);
                    }
                    else
                    {
                        _logger.Info("Case workbook hidden session workbook close starting. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                        WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave(
                            workbook,
                            _logger,
                            "CaseWorkbookOpenStrategy.CleanupCachedHiddenSession route=" + (routeName ?? string.Empty));
                        workbookCloseCompleted = true;
                        _logger.Info("Case workbook hidden session workbook close completed. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
                    }
                }
            }
            catch (Exception ex)
            {
                closeFailed = true;
                closeFailureType = ex.GetType().Name;
                markPoisoned = true;
                _logger.Error("CleanupCachedHiddenSession workbook close failed.", ex);
	            }
	            finally
	            {
	                Stopwatch stopwatch3 = Stopwatch.StartNew();
	                ReleaseComObject(workbook);
	                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "waitUiShownToCaseCreated", "comRelease", stopwatch3.ElapsedMilliseconds, "route=" + (routeName ?? string.Empty));
	            }

	            if (markPoisoned || closeFailed)
	            {
	                MarkCachedHiddenApplicationPoisoned(
                        application,
                        caseWorkbookPath,
                        routeName,
                        stopwatch,
                        closeFailed ? "workbookCloseFailed" : "markPoisoned",
                        closeFailureType);
                    LogHiddenCleanupOutcome(_cleanupOutcomeService.CreateCachedHiddenSessionPoisonedOutcome(
                        "CaseWorkbookOpenStrategy.CleanupCachedHiddenSession",
                        caseWorkbookPath,
                        routeName,
                        _hiddenAppLifecycleSupportService.CreateCachedSessionCleanupFacts(
                            workbookPresent,
                            workbookCloseAttempted,
                            workbookCloseCompleted,
                            application != null),
                        closeFailed ? "workbookCloseFailed" : "markPoisoned"));
	                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "waitUiShownToCaseCreated", "hiddenSessionClose", stopwatch2.ElapsedMilliseconds, "route=" + (routeName ?? string.Empty) + ", cached=False");
	                _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", cached=False, elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
	                return;
	            }

	            if (TryReturnCachedHiddenApplicationToIdle(application, caseWorkbookPath, routeName, stopwatch))
	            {
                    LogHiddenCleanupOutcome(_cleanupOutcomeService.CreateCachedHiddenSessionReturnedToIdleOutcome(
                        "CaseWorkbookOpenStrategy.CleanupCachedHiddenSession",
                        caseWorkbookPath,
                        routeName,
                        _hiddenAppLifecycleSupportService.CreateCachedSessionCleanupFacts(
                            workbookPresent,
                            workbookCloseAttempted,
                            workbookCloseCompleted,
                            application != null)));
	                NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "waitUiShownToCaseCreated", "hiddenSessionClose", stopwatch2.ElapsedMilliseconds, "route=" + (routeName ?? string.Empty) + ", cached=True");
	                _logger.Info("Case workbook hidden session close finalized. path=" + (caseWorkbookPath ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", cached=True, elapsedMs=" + ((stopwatch == null) ? string.Empty : stopwatch.ElapsedMilliseconds.ToString()));
	                return;
	            }

	            MarkCachedHiddenApplicationPoisoned(
                    application,
                    caseWorkbookPath,
                    routeName,
                    stopwatch,
                    "returnToIdleFailed",
                    string.Empty);
                LogHiddenCleanupOutcome(_cleanupOutcomeService.CreateCachedHiddenSessionPoisonedOutcome(
                    "CaseWorkbookOpenStrategy.CleanupCachedHiddenSession",
                    caseWorkbookPath,
                    routeName,
                    _hiddenAppLifecycleSupportService.CreateCachedSessionCleanupFacts(
                        workbookPresent,
                        workbookCloseAttempted,
                        workbookCloseCompleted,
                        application != null),
                    "returnToIdleFailed"));
	            NewCaseDefaultTimingLogHelper.LogDetail(_logger, caseWorkbookPath, "waitUiShownToCaseCreated", "hiddenSessionClose", stopwatch2.ElapsedMilliseconds, "route=" + (routeName ?? string.Empty) + ", cached=False");
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
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionIdleReturn,
                        caseWorkbookPath,
                        routeName,
                        "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                        "cache-slot-mismatch")
                    {
                        EventOutcome = "not-returned",
                        CacheEvent = "idle-return",
                        ReturnOutcome = "not-cache-slot",
                        AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(application),
                        RetainedInstancePresent = _cachedHiddenApplication != null && _cachedHiddenApplication.Application != null,
                        SafetyAction = "return-false",
                        ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                    });
                    return false;
                }

                if (!_routeDecisionService.IsHiddenApplicationCacheEnabled())
                {
                    _logger.Info(_hiddenAppLifecycleSupportService.BuildReturnToIdleDisabledMessage(caseWorkbookPath, routeName, _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(application)));
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionIdleReturn,
                        caseWorkbookPath,
                        routeName,
                        "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                        "feature-flag-disabled")
                    {
                        EventOutcome = "not-returned",
                        CacheEvent = "idle-return",
                        ReturnOutcome = "discard",
                        AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(application),
                        SafetyAction = "poison-dispose",
                        ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                    });
                    _cachedHiddenApplication.IsPoisoned = true;
                    return false;
                }

                try
                {
                    PrepareHiddenApplicationForUse(application);
                }
                catch (Exception ex)
                {
                    _logger.Error(_hiddenAppLifecycleSupportService.BuildReturnToIdleFailedHiddenStateMessage(), ex);
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionIdleReturn,
                        caseWorkbookPath,
                        routeName,
                        "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                        "hidden-state-reapply-failed")
                    {
                        EventOutcome = "not-returned",
                        CacheEvent = "idle-return",
                        ReturnOutcome = "discard",
                        AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(application),
                        ExceptionType = ex.GetType().Name,
                        SafetyAction = "poison-dispose",
                        ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                    });
                    _cachedHiddenApplication.IsPoisoned = true;
                    return false;
                }

                CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts = CaptureCachedHiddenApplicationLifecycleFactsUnlocked(_cachedHiddenApplication);
                if (!lifecycleFacts.IsApplicationStateHealthy)
                {
                    _cachedHiddenApplication.IsPoisoned = true;
                    _logger.Warn(_hiddenAppLifecycleSupportService.BuildCacheUnhealthyMessage("return-to-idle-health-check-failed", lifecycleFacts.AppHwnd));
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionOrphanSuspicion,
                        caseWorkbookPath,
                        routeName,
                        "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                        "return-to-idle-health-check-failed")
                    {
                        EventOutcome = "suspected-unhealthy-retained-instance",
                        CacheEvent = "idle-return-health-check",
                        AppHwnd = lifecycleFacts.AppHwnd,
                        SafetyAction = "poison-dispose",
                        AbandonedOperation = "return-retained-application-to-idle",
                        LifecycleFacts = lifecycleFacts,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                    });
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionIdleReturn,
                        caseWorkbookPath,
                        routeName,
                        "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                        "health-check-failed")
                    {
                        EventOutcome = "not-returned",
                        CacheEvent = "idle-return",
                        ReturnOutcome = "discard",
                        AppHwnd = lifecycleFacts.AppHwnd,
                        SafetyAction = "poison-dispose",
                        LifecycleFacts = lifecycleFacts,
                        ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                    });
                    return false;
                }

                _cachedHiddenApplication.IsInUse = false;
                _cachedHiddenApplication.IdleSinceUtc = DateTime.UtcNow;
                ScheduleHiddenApplicationIdleTimerUnlocked();
                _logger.Info(_hiddenAppLifecycleSupportService.BuildReturnedToIdleMessage(caseWorkbookPath, routeName, lifecycleFacts.AppHwnd, _routeDecisionService.ResolveHiddenApplicationCacheIdleSeconds(), stopwatch == null ? 0 : stopwatch.ElapsedMilliseconds));
                LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                    CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionIdleReturn,
                    caseWorkbookPath,
                    routeName,
                    "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                    "returnedToIdle")
                {
                    EventOutcome = "returned-to-idle",
                    CacheEvent = "idle-return",
                    ReturnOutcome = "returned-to-idle",
                    AppHwnd = lifecycleFacts.AppHwnd,
                    SafetyAction = "keep-retained-application-idle",
                    LifecycleFacts = lifecycleFacts,
                    ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                    ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                });
                NewCaseVisibilityObservation.Log(
                    _logger,
                    null,
                    application,
                    null,
                    null,
                    "post-create-cleanup",
                    "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                    caseWorkbookPath,
                    _hiddenAppLifecycleSupportService.BuildPostCreateCleanupObservationDetails(routeName));
                return true;
            }
        }

        private void MarkCachedHiddenApplicationPoisoned(
            Excel.Application application,
            string caseWorkbookPath,
            string routeName,
            Stopwatch stopwatch,
            string poisonReason,
            string exceptionType)
        {
            CachedHiddenApplicationSlot slotToDispose = null;
            string appHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(application);
            lock (_hiddenApplicationCacheSync)
            {
                if (_cachedHiddenApplication != null && ReferenceEquals(_cachedHiddenApplication.Application, application))
                {
                    _cachedHiddenApplication.IsPoisoned = true;
                    _logger.Warn(_hiddenAppLifecycleSupportService.BuildPoisonedMessage(caseWorkbookPath, routeName, appHwnd, stopwatch == null ? 0 : stopwatch.ElapsedMilliseconds, poisonReason));
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionPoisonMark,
                        caseWorkbookPath,
                        routeName,
                        "CaseWorkbookOpenStrategy.MarkCachedHiddenApplicationPoisoned",
                        poisonReason)
                    {
                        EventOutcome = "marked",
                        CacheEvent = "poison",
                        PoisonReason = poisonReason,
                        AppHwnd = appHwnd,
                        ExceptionType = exceptionType,
                        SafetyAction = "detach-and-dispose-retained-application",
                        ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                    });
                    slotToDispose = _cachedHiddenApplication;
                    _cachedHiddenApplication = null;
                    StopHiddenApplicationIdleTimerUnlocked();
                }
            }

            if (slotToDispose == null)
            {
                LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                    CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionPoisonMark,
                    caseWorkbookPath,
                    routeName,
                    "CaseWorkbookOpenStrategy.MarkCachedHiddenApplicationPoisoned",
                    poisonReason)
                {
                    EventOutcome = "not-marked-slot-missing",
                    CacheEvent = "poison",
                    PoisonReason = poisonReason,
                    AppHwnd = appHwnd,
                    ExceptionType = exceptionType,
                    SafetyAction = "skip-dispose-slot-not-found",
                    ElapsedMilliseconds = stopwatch == null ? (long?)null : stopwatch.ElapsedMilliseconds,
                    ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName)
                });
            }

            DisposeCachedHiddenApplicationSlot(slotToDispose, "poisoned");
        }

        private bool TryCloseWorkbookWithoutSaving(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave(
                    workbook,
                    _logger,
                    "CaseWorkbookOpenStrategy.TryCloseWorkbookWithoutSaving");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("TryCloseWorkbookWithoutSaving failed.", ex);
                return false;
            }
        }

        private bool TryQuitApplication(Excel.Application application)
        {
            if (application == null)
            {
                return false;
            }

            try
            {
                application.Quit();
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("TryQuitApplication failed.", ex);
                return false;
            }
        }

        private void LogHiddenCleanupOutcome(CaseWorkbookOpenHiddenCleanupOutcome outcome)
        {
            if (outcome == null)
            {
                return;
            }

            _logger.Info(outcome.KernelFlickerTraceMessage);
            NewCaseVisibilityObservation.Log(
                _logger,
                null,
                null,
                null,
                null,
                "hidden-excel-cleanup-outcome",
                outcome.Owner,
                outcome.CaseWorkbookPath,
                outcome.Details);
        }

        private void LogRetainedInstanceCleanupOutcome(CaseWorkbookOpenRetainedCleanupOutcome outcome)
        {
            if (outcome == null)
            {
                return;
            }

            _logger.Info(outcome.KernelFlickerTraceMessage);
        }

        private void LogRetainedHiddenAppCacheEvent(CaseWorkbookHiddenAppLifecycleDiagnosticEvent diagnosticEvent)
        {
            _logger.Info(_hiddenAppLifecycleSupportService.BuildDiagnosticEventMessage(diagnosticEvent));
        }

        private void HiddenApplicationIdleTimer_Tick(object sender, EventArgs e)
        {
            if (!_routeDecisionService.IsHiddenApplicationCacheEnabled())
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
            CaseWorkbookHiddenAppExpirationDecision expirationDecision =
                _hiddenAppLifecycleSupportService.DecideExpiration(
                    CreateCachedHiddenApplicationExpirationFactsUnlocked(_cachedHiddenApplication),
                    reason);
            if (expirationDecision.InitializeIdleSinceUtc && _cachedHiddenApplication != null)
            {
                _cachedHiddenApplication.IdleSinceUtc = expirationDecision.InitializedIdleSinceUtc;
            }

            if (!expirationDecision.DisposeSlot)
            {
                if (string.Equals(expirationDecision.DecisionReason, "in-use", StringComparison.Ordinal))
                {
                    LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                        CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionTimeoutFallback,
                        string.Empty,
                        CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                        "CaseWorkbookOpenStrategy.CleanupExpiredCachedHiddenApplicationUnlocked",
                        expirationDecision.DecisionReason)
                    {
                        EventOutcome = "fallback-stop-idle-timer",
                        CacheEvent = "expiration-decision",
                        CleanupReason = reason,
                        AbandonedOperation = "idle-timeout-cleanup",
                        SafetyAction = "stop-idle-timer",
                        ExpirationDecision = expirationDecision,
                        ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
                    });
                }

                if (expirationDecision.StopIdleTimer)
                {
                    StopHiddenApplicationIdleTimerUnlocked();
                }

                return null;
            }

            CachedHiddenApplicationSlot expiredSlot = _cachedHiddenApplication;
            _cachedHiddenApplication = null;
            if (expirationDecision.StopIdleTimer)
            {
                StopHiddenApplicationIdleTimerUnlocked();
            }

            if (expiredSlot != null && !expiredSlot.IsPoisoned)
            {
                _logger.Info(_hiddenAppLifecycleSupportService.BuildTimedOutMessage(reason, _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(expiredSlot.Application)));
            }

            LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionTimeoutFallback,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.CleanupExpiredCachedHiddenApplicationUnlocked",
                expirationDecision.DecisionReason)
            {
                EventOutcome = "dispose-retained-instance",
                CacheEvent = "expiration-decision",
                CleanupReason = reason,
                AppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(expiredSlot == null ? null : expiredSlot.Application),
                AbandonedOperation = "retain-idle-cache",
                SafetyAction = "dispose-retained-application",
                ExpirationDecision = expirationDecision,
                ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            });
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
                string appHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(slot.Application);
                _logger.Warn(_hiddenAppLifecycleSupportService.BuildCleanupSkippedNotOwnedMessage(reason, appHwnd));
                LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                    CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionDispose,
                    string.Empty,
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                    "CaseWorkbookOpenStrategy.DisposeCachedHiddenApplicationSlot",
                    reason)
                {
                    EventOutcome = "skipped-not-cache-owned",
                    CacheEvent = "dispose",
                    CleanupReason = reason,
                    AppHwnd = appHwnd,
                    RetainedInstancePresent = slot.Application != null,
                    AppQuitAttempted = false,
                    AppQuitCompleted = false,
                    SafetyAction = "skip-quit",
                    ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
                });
                LogRetainedInstanceCleanupOutcome(_cleanupOutcomeService.CreateRetainedInstanceCleanupOutcome(
                    reason,
                    appHwnd,
                    retainedInstancePresent: slot.Application != null,
                    isOwnedByCache: false,
                    quitAttempted: false,
                    quitCompleted: false));
                return;
            }

            bool quitCompleted = TryQuitApplication(slot.Application);
            ReleaseComObject(slot.Application);
            string retainedAppHwnd = _hiddenAppLifecycleSupportService.CaptureApplicationHwnd(slot.Application);
            LogRetainedInstanceCleanupOutcome(_cleanupOutcomeService.CreateRetainedInstanceCleanupOutcome(
                reason,
                retainedAppHwnd,
                retainedInstancePresent: slot.Application != null,
                isOwnedByCache: true,
                quitAttempted: slot.Application != null,
                quitCompleted: quitCompleted));
            LogRetainedHiddenAppCacheEvent(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionDispose,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.DisposeCachedHiddenApplicationSlot",
                reason)
            {
                EventOutcome = quitCompleted ? "disposed" : "degraded",
                CacheEvent = "dispose",
                CleanupReason = reason,
                AppHwnd = retainedAppHwnd,
                RetainedInstancePresent = slot.Application != null,
                AppQuitAttempted = slot.Application != null,
                AppQuitCompleted = quitCompleted,
                SafetyAction = quitCompleted ? "quit-and-release" : "release-after-quit-failure",
                ApplicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            });
            _logger.Info(_hiddenAppLifecycleSupportService.BuildDiscardedMessage(reason, retainedAppHwnd));
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

        private CaseWorkbookHiddenAppLifecycleFacts CaptureCachedHiddenApplicationLifecycleFactsUnlocked(CachedHiddenApplicationSlot slot)
        {
            return _hiddenAppLifecycleSupportService.CaptureLifecycleFacts(
                slot == null ? null : slot.Application,
                slot != null && slot.IsInUse,
                slot != null && slot.IsPoisoned,
                slot != null && slot.IsOwnedByCache,
                slot == null ? DateTime.MinValue : slot.IdleSinceUtc,
                _routeDecisionService.ResolveHiddenApplicationCacheIdleSeconds(),
                DateTime.UtcNow);
        }

        private CaseWorkbookHiddenAppLifecycleFacts CreateCachedHiddenApplicationExpirationFactsUnlocked(CachedHiddenApplicationSlot slot)
        {
            return _hiddenAppLifecycleSupportService.CreateLifecycleStateFacts(
                applicationPresent: slot != null && slot.Application != null,
                isInUse: slot != null && slot.IsInUse,
                isPoisoned: slot != null && slot.IsPoisoned,
                isOwnedByCache: slot != null && slot.IsOwnedByCache,
                appHwnd: string.Empty,
                idleSinceUtc: slot == null ? DateTime.MinValue : slot.IdleSinceUtc,
                idleTimeoutSeconds: _routeDecisionService.ResolveHiddenApplicationCacheIdleSeconds(),
                utcNow: DateTime.UtcNow);
        }

        private void ReleaseComObject(object comObject, [CallerMemberName] string callerMemberName = null)
        {
            if (comObject == null)
            {
                return;
            }

            if (_releaseComObject != null)
            {
                try
                {
                    _releaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    _logger.Error("ReleaseComObject hook failed.", ex);
                }
            }

            // Hidden Excel セッション由来の所有参照は完全解放の方針を維持する。
            ComObjectReleaseService.FinalRelease(
                comObject,
                _logger,
                nameof(CaseWorkbookOpenStrategy) + "." + (callerMemberName ?? nameof(ReleaseComObject)));
        }

        private void HideOpenedWorkbookWindow(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            Excel.Windows windows = null;
            try
            {
                windows = workbook.Windows;
                int windowCount = windows == null ? 0 : windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = windows[index];
                        if (window != null)
                        {
                            window.Visible = false;
                        }
                    }
                    finally
                    {
                        ComObjectReleaseService.Release(window, _logger);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("HideOpenedWorkbookWindow failed.", ex);
            }
            finally
            {
                ComObjectReleaseService.Release(windows, _logger);
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
