using System;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookHiddenAppLifecycleSupportService
    {
        internal const string ReuseBlockReasonApplicationMissing = "application-missing";
        internal const string ReuseBlockReasonInUse = "in-use";
        internal const string ReuseBlockReasonPoisoned = "poisoned";
        internal const string ReuseBlockReasonFactsUnavailable = "application-facts-unavailable";
        internal const string ReuseBlockReasonWorkbooksUnavailable = "workbooks-unavailable";
        internal const string ReuseBlockReasonWorkbooksOpen = "workbooks-open";
        internal const string ReuseBlockReasonNotReady = "not-ready";
        internal const string ReuseBlockReasonApplicationVisible = "application-visible";
        internal const string ReuseBlockReasonDisplayAlertsEnabled = "display-alerts-enabled";
        internal const string ReuseBlockReasonScreenUpdatingEnabled = "screen-updating-enabled";
        internal const string ReuseBlockReasonEventsEnabled = "events-enabled";
        internal const string ReuseBlockReasonUserControlEnabled = "user-control-enabled";

        internal CaseWorkbookHiddenAppLifecycleFacts CaptureLifecycleFacts(
            Excel.Application application,
            bool isInUse,
            bool isPoisoned,
            bool isOwnedByCache,
            DateTime idleSinceUtc,
            int idleTimeoutSeconds,
            DateTime utcNow)
        {
            string appHwnd = CaptureApplicationHwnd(application);
            if (application == null)
            {
                return new CaseWorkbookHiddenAppLifecycleFacts(
                    applicationPresent: false,
                    isInUse: isInUse,
                    isPoisoned: isPoisoned,
                    isOwnedByCache: isOwnedByCache,
                    appHwnd: appHwnd,
                    captureSucceeded: true,
                    captureFailureType: string.Empty,
                    workbooksPresent: false,
                    workbooksCount: -1,
                    ready: false,
                    visible: false,
                    displayAlerts: false,
                    screenUpdating: false,
                    enableEvents: false,
                    userControl: false,
                    idleSinceUtc: idleSinceUtc,
                    idleTimeoutSeconds: idleTimeoutSeconds,
                    utcNow: utcNow);
            }

            try
            {
                Excel.Workbooks workbooks = application.Workbooks;
                bool workbooksPresent = workbooks != null;
                int workbooksCount = workbooksPresent ? workbooks.Count : -1;
                return new CaseWorkbookHiddenAppLifecycleFacts(
                    applicationPresent: true,
                    isInUse: isInUse,
                    isPoisoned: isPoisoned,
                    isOwnedByCache: isOwnedByCache,
                    appHwnd: appHwnd,
                    captureSucceeded: true,
                    captureFailureType: string.Empty,
                    workbooksPresent: workbooksPresent,
                    workbooksCount: workbooksCount,
                    ready: application.Ready,
                    visible: application.Visible,
                    displayAlerts: application.DisplayAlerts,
                    screenUpdating: application.ScreenUpdating,
                    enableEvents: application.EnableEvents,
                    userControl: application.UserControl,
                    idleSinceUtc: idleSinceUtc,
                    idleTimeoutSeconds: idleTimeoutSeconds,
                    utcNow: utcNow);
            }
            catch (Exception ex)
            {
                return new CaseWorkbookHiddenAppLifecycleFacts(
                    applicationPresent: true,
                    isInUse: isInUse,
                    isPoisoned: isPoisoned,
                    isOwnedByCache: isOwnedByCache,
                    appHwnd: appHwnd,
                    captureSucceeded: false,
                    captureFailureType: ex.GetType().Name,
                    workbooksPresent: false,
                    workbooksCount: -1,
                    ready: false,
                    visible: false,
                    displayAlerts: false,
                    screenUpdating: false,
                    enableEvents: false,
                    userControl: false,
                    idleSinceUtc: idleSinceUtc,
                    idleTimeoutSeconds: idleTimeoutSeconds,
                    utcNow: utcNow);
            }
        }

        internal CaseWorkbookHiddenAppLifecycleFacts CreateLifecycleStateFacts(
            bool applicationPresent,
            bool isInUse,
            bool isPoisoned,
            bool isOwnedByCache,
            string appHwnd,
            DateTime idleSinceUtc,
            int idleTimeoutSeconds,
            DateTime utcNow)
        {
            return new CaseWorkbookHiddenAppLifecycleFacts(
                applicationPresent: applicationPresent,
                isInUse: isInUse,
                isPoisoned: isPoisoned,
                isOwnedByCache: isOwnedByCache,
                appHwnd: appHwnd,
                captureSucceeded: true,
                captureFailureType: string.Empty,
                workbooksPresent: false,
                workbooksCount: -1,
                ready: false,
                visible: false,
                displayAlerts: false,
                screenUpdating: false,
                enableEvents: false,
                userControl: false,
                idleSinceUtc: idleSinceUtc,
                idleTimeoutSeconds: idleTimeoutSeconds,
                utcNow: utcNow);
        }

        internal CaseWorkbookHiddenAppExpirationDecision DecideExpiration(
            CaseWorkbookHiddenAppLifecycleFacts facts,
            string cleanupReason)
        {
            CaseWorkbookHiddenAppLifecycleFacts safeFacts = facts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            if (!safeFacts.ApplicationPresent)
            {
                return new CaseWorkbookHiddenAppExpirationDecision(
                    disposeSlot: false,
                    stopIdleTimer: true,
                    initializeIdleSinceUtc: false,
                    initializedIdleSinceUtc: safeFacts.IdleSinceUtc,
                    decisionReason: "cache-empty",
                    cleanupReason: cleanupReason,
                    facts: safeFacts);
            }

            if (safeFacts.IsInUse)
            {
                return new CaseWorkbookHiddenAppExpirationDecision(
                    disposeSlot: false,
                    stopIdleTimer: true,
                    initializeIdleSinceUtc: false,
                    initializedIdleSinceUtc: safeFacts.IdleSinceUtc,
                    decisionReason: "in-use",
                    cleanupReason: cleanupReason,
                    facts: safeFacts);
            }

            if (safeFacts.IsPoisoned)
            {
                return new CaseWorkbookHiddenAppExpirationDecision(
                    disposeSlot: true,
                    stopIdleTimer: true,
                    initializeIdleSinceUtc: false,
                    initializedIdleSinceUtc: safeFacts.IdleSinceUtc,
                    decisionReason: "poisoned",
                    cleanupReason: cleanupReason,
                    facts: safeFacts);
            }

            if (safeFacts.IdleSinceUtc == DateTime.MinValue)
            {
                return new CaseWorkbookHiddenAppExpirationDecision(
                    disposeSlot: false,
                    stopIdleTimer: false,
                    initializeIdleSinceUtc: true,
                    initializedIdleSinceUtc: safeFacts.UtcNow,
                    decisionReason: "idle-since-uninitialized",
                    cleanupReason: cleanupReason,
                    facts: safeFacts);
            }

            if (!safeFacts.IdleTimeoutExpired)
            {
                return new CaseWorkbookHiddenAppExpirationDecision(
                    disposeSlot: false,
                    stopIdleTimer: false,
                    initializeIdleSinceUtc: false,
                    initializedIdleSinceUtc: safeFacts.IdleSinceUtc,
                    decisionReason: "idle-timeout-not-reached",
                    cleanupReason: cleanupReason,
                    facts: safeFacts);
            }

            return new CaseWorkbookHiddenAppExpirationDecision(
                disposeSlot: true,
                stopIdleTimer: true,
                initializeIdleSinceUtc: false,
                initializedIdleSinceUtc: safeFacts.IdleSinceUtc,
                decisionReason: "idle-timeout-reached",
                cleanupReason: cleanupReason,
                facts: safeFacts);
        }

        internal CaseWorkbookOpenCleanupFacts CreateCachedSessionCleanupFacts(
            bool workbookPresent,
            bool workbookCloseAttempted,
            bool workbookCloseCompleted,
            bool appPresent)
        {
            return new CaseWorkbookOpenCleanupFacts(
                workbookPresent,
                workbookCloseAttempted,
                workbookCloseCompleted,
                appPresent,
                appQuitAttempted: false,
                appQuitCompleted: false);
        }

        internal string CaptureApplicationHwnd(Excel.Application application)
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

        internal string BuildCacheBypassInUseMessage(
            string caseWorkbookPath,
            CaseWorkbookOpenRouteDecision fallbackDecision,
            long elapsedMilliseconds)
        {
            string routeName = fallbackDecision == null
                ? CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheBypassInUseRouteName
                : fallbackDecision.RouteName;
            string routeReason = fallbackDecision == null
                ? "hiddenApplicationCacheInUse"
                : fallbackDecision.Reason;
            return "hidden-app-cache bypassed because in-use. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + routeName
                + ", routeReason="
                + routeReason
                + ", elapsedMs="
                + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture);
        }

        internal string BuildCacheAcquiredMessage(
            string caseWorkbookPath,
            bool reusedApplication,
            string routeName,
            string appHwnd,
            long elapsedMilliseconds)
        {
            return "hidden-app-cache "
                + (reusedApplication ? "reused" : "created")
                + ". path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + (routeName ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty)
                + ", elapsedMs="
                + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture);
        }

        internal string BuildCacheUnhealthyMessage(string reason, string appHwnd)
        {
            return "hidden-app-cache unhealthy. reason="
                + (reason ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty);
        }

        internal string BuildReturnToIdleDisabledMessage(
            string caseWorkbookPath,
            string routeName,
            string appHwnd)
        {
            return "hidden-app-cache disabled before return-to-idle. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + (routeName ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty);
        }

        internal string BuildReturnToIdleFailedHiddenStateMessage()
        {
            return "hidden-app-cache failed to reapply hidden state.";
        }

        internal string BuildReturnedToIdleMessage(
            string caseWorkbookPath,
            string routeName,
            string appHwnd,
            int idleTimeoutSeconds,
            long elapsedMilliseconds)
        {
            return "hidden-app-cache returned-to-idle. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + (routeName ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty)
                + ", idleTimeoutSeconds="
                + idleTimeoutSeconds.ToString(CultureInfo.InvariantCulture)
                + ", elapsedMs="
                + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture);
        }

        internal string BuildPoisonedMessage(
            string caseWorkbookPath,
            string routeName,
            string appHwnd,
            long elapsedMilliseconds)
        {
            return "hidden-app-cache poisoned. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + (routeName ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty)
                + ", elapsedMs="
                + elapsedMilliseconds.ToString(CultureInfo.InvariantCulture);
        }

        internal string BuildTimedOutMessage(string reason, string appHwnd)
        {
            return "hidden-app-cache timed-out. reason="
                + (reason ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty);
        }

        internal string BuildCleanupSkippedNotOwnedMessage(string reason, string appHwnd)
        {
            return "hidden-app-cache cleanup skipped because slot is not cache-owned. reason="
                + (reason ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty);
        }

        internal string BuildDiscardedMessage(string reason, string appHwnd)
        {
            return "hidden-app-cache discarded. reason="
                + (reason ?? string.Empty)
                + ", appHwnd="
                + (appHwnd ?? string.Empty);
        }

        internal string BuildAcquiredObservationDetails(
            string routeName,
            bool reusedApplication,
            string applicationOwnerFacts)
        {
            return "route=" + (routeName ?? string.Empty)
                + ",reused=" + reusedApplication.ToString()
                + "," + (applicationOwnerFacts ?? string.Empty);
        }

        internal string BuildPostCreateCleanupObservationDetails(string routeName)
        {
            return "route=" + (routeName ?? string.Empty);
        }
    }

    internal sealed class CaseWorkbookHiddenAppLifecycleFacts
    {
        internal static readonly CaseWorkbookHiddenAppLifecycleFacts Empty =
            new CaseWorkbookHiddenAppLifecycleFacts(
                applicationPresent: false,
                isInUse: false,
                isPoisoned: false,
                isOwnedByCache: false,
                appHwnd: string.Empty,
                captureSucceeded: true,
                captureFailureType: string.Empty,
                workbooksPresent: false,
                workbooksCount: -1,
                ready: false,
                visible: false,
                displayAlerts: false,
                screenUpdating: false,
                enableEvents: false,
                userControl: false,
                idleSinceUtc: DateTime.MinValue,
                idleTimeoutSeconds: 0,
                utcNow: DateTime.MinValue);

        internal CaseWorkbookHiddenAppLifecycleFacts(
            bool applicationPresent,
            bool isInUse,
            bool isPoisoned,
            bool isOwnedByCache,
            string appHwnd,
            bool captureSucceeded,
            string captureFailureType,
            bool workbooksPresent,
            int workbooksCount,
            bool ready,
            bool visible,
            bool displayAlerts,
            bool screenUpdating,
            bool enableEvents,
            bool userControl,
            DateTime idleSinceUtc,
            int idleTimeoutSeconds,
            DateTime utcNow)
        {
            ApplicationPresent = applicationPresent;
            IsInUse = isInUse;
            IsPoisoned = isPoisoned;
            IsOwnedByCache = isOwnedByCache;
            AppHwnd = appHwnd ?? string.Empty;
            CaptureSucceeded = captureSucceeded;
            CaptureFailureType = captureFailureType ?? string.Empty;
            WorkbooksPresent = workbooksPresent;
            WorkbooksCount = workbooksCount;
            Ready = ready;
            Visible = visible;
            DisplayAlerts = displayAlerts;
            ScreenUpdating = screenUpdating;
            EnableEvents = enableEvents;
            UserControl = userControl;
            IdleSinceUtc = idleSinceUtc;
            IdleTimeoutSeconds = idleTimeoutSeconds;
            UtcNow = utcNow;
        }

        internal bool ApplicationPresent { get; }

        internal bool IsInUse { get; }

        internal bool IsPoisoned { get; }

        internal bool IsOwnedByCache { get; }

        internal string AppHwnd { get; }

        internal bool CaptureSucceeded { get; }

        internal string CaptureFailureType { get; }

        internal bool WorkbooksPresent { get; }

        internal int WorkbooksCount { get; }

        internal bool Ready { get; }

        internal bool Visible { get; }

        internal bool DisplayAlerts { get; }

        internal bool ScreenUpdating { get; }

        internal bool EnableEvents { get; }

        internal bool UserControl { get; }

        internal DateTime IdleSinceUtc { get; }

        internal int IdleTimeoutSeconds { get; }

        internal DateTime UtcNow { get; }

        internal double IdleAgeSeconds
        {
            get
            {
                if (IdleSinceUtc == DateTime.MinValue)
                {
                    return 0;
                }

                return (UtcNow - IdleSinceUtc).TotalSeconds;
            }
        }

        internal bool IdleTimeoutExpired
        {
            get
            {
                return IdleSinceUtc != DateTime.MinValue
                    && IdleTimeoutSeconds >= 0
                    && IdleAgeSeconds >= IdleTimeoutSeconds;
            }
        }

        internal bool IsReusable
        {
            get
            {
                return !IsInUse && IsApplicationStateHealthy;
            }
        }

        internal string ReuseBlockReason
        {
            get
            {
                if (!ApplicationPresent)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonApplicationMissing;
                }

                if (IsInUse)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonInUse;
                }

                return ApplicationStateBlockReason;
            }
        }

        internal bool IsApplicationStateHealthy
        {
            get
            {
                return string.IsNullOrEmpty(ApplicationStateBlockReason);
            }
        }

        internal string ApplicationStateBlockReason
        {
            get
            {
                if (!ApplicationPresent)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonApplicationMissing;
                }

                if (IsPoisoned)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonPoisoned;
                }

                if (!CaptureSucceeded)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonFactsUnavailable;
                }

                if (!WorkbooksPresent)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonWorkbooksUnavailable;
                }

                if (WorkbooksCount != 0)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonWorkbooksOpen;
                }

                if (!Ready)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonNotReady;
                }

                if (Visible)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonApplicationVisible;
                }

                if (DisplayAlerts)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonDisplayAlertsEnabled;
                }

                if (ScreenUpdating)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonScreenUpdatingEnabled;
                }

                if (EnableEvents)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonEventsEnabled;
                }

                if (UserControl)
                {
                    return CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonUserControlEnabled;
                }

                return string.Empty;
            }
        }

        internal string DiagnosticDetails
        {
            get
            {
                return "scope=hidden-app-lifecycle"
                    + ",appPresent=" + ApplicationPresent.ToString()
                    + ",isInUse=" + IsInUse.ToString()
                    + ",isPoisoned=" + IsPoisoned.ToString()
                    + ",isOwnedByCache=" + IsOwnedByCache.ToString()
                    + ",appHwnd=" + AppHwnd
                    + ",captureSucceeded=" + CaptureSucceeded.ToString()
                    + ",captureFailureType=" + CaptureFailureType
                    + ",workbooksPresent=" + WorkbooksPresent.ToString()
                    + ",workbooksCount=" + WorkbooksCount.ToString(CultureInfo.InvariantCulture)
                    + ",ready=" + Ready.ToString()
                    + ",visible=" + Visible.ToString()
                    + ",displayAlerts=" + DisplayAlerts.ToString()
                    + ",screenUpdating=" + ScreenUpdating.ToString()
                    + ",enableEvents=" + EnableEvents.ToString()
                    + ",userControl=" + UserControl.ToString()
                    + ",idleSinceUtc=" + FormatDateTime(IdleSinceUtc)
                    + ",idleTimeoutSeconds=" + IdleTimeoutSeconds.ToString(CultureInfo.InvariantCulture)
                    + ",idleAgeSeconds=" + IdleAgeSeconds.ToString("0.###", CultureInfo.InvariantCulture)
                    + ",idleTimeoutExpired=" + IdleTimeoutExpired.ToString()
                    + ",isReusable=" + IsReusable.ToString()
                    + ",reuseBlockReason=" + ReuseBlockReason
                    + ",isApplicationStateHealthy=" + IsApplicationStateHealthy.ToString()
                    + ",applicationStateBlockReason=" + ApplicationStateBlockReason;
            }
        }

        private static string FormatDateTime(DateTime value)
        {
            return value == DateTime.MinValue ? string.Empty : value.ToString("o", CultureInfo.InvariantCulture);
        }
    }

    internal sealed class CaseWorkbookHiddenAppExpirationDecision
    {
        internal CaseWorkbookHiddenAppExpirationDecision(
            bool disposeSlot,
            bool stopIdleTimer,
            bool initializeIdleSinceUtc,
            DateTime initializedIdleSinceUtc,
            string decisionReason,
            string cleanupReason,
            CaseWorkbookHiddenAppLifecycleFacts facts)
        {
            DisposeSlot = disposeSlot;
            StopIdleTimer = stopIdleTimer;
            InitializeIdleSinceUtc = initializeIdleSinceUtc;
            InitializedIdleSinceUtc = initializedIdleSinceUtc;
            DecisionReason = decisionReason ?? string.Empty;
            CleanupReason = cleanupReason ?? string.Empty;
            Facts = facts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
        }

        internal bool DisposeSlot { get; }

        internal bool StopIdleTimer { get; }

        internal bool InitializeIdleSinceUtc { get; }

        internal DateTime InitializedIdleSinceUtc { get; }

        internal string DecisionReason { get; }

        internal string CleanupReason { get; }

        internal CaseWorkbookHiddenAppLifecycleFacts Facts { get; }

        internal string DiagnosticDetails
        {
            get
            {
                return "scope=hidden-app-expiration"
                    + ",cleanupReason=" + CleanupReason
                    + ",decisionReason=" + DecisionReason
                    + ",disposeSlot=" + DisposeSlot.ToString()
                    + ",stopIdleTimer=" + StopIdleTimer.ToString()
                    + ",initializeIdleSinceUtc=" + InitializeIdleSinceUtc.ToString()
                    + "," + Facts.DiagnosticDetails;
            }
        }
    }
}
