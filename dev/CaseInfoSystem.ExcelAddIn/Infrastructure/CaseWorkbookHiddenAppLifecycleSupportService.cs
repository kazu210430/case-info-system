using System;
using System.Globalization;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookHiddenAppLifecycleSupportService
    {
        internal const string LifecycleActionAcquire = "retained-hidden-app-cache-acquire";
        internal const string LifecycleActionIdleReturn = "retained-hidden-app-cache-idle-return";
        internal const string LifecycleActionPoisonMark = "retained-hidden-app-cache-poison-mark";
        internal const string LifecycleActionShutdown = "retained-hidden-app-cache-shutdown";
        internal const string LifecycleActionDispose = "retained-hidden-app-cache-dispose";
        internal const string LifecycleActionTimeoutFallback = "retained-hidden-app-cache-timeout-fallback";
        internal const string LifecycleActionFallback = "retained-hidden-app-cache-fallback";
        internal const string LifecycleActionOrphanSuspicion = "retained-hidden-app-cache-orphan-suspicion";

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
            long elapsedMilliseconds,
            string poisonReason)
        {
            return "hidden-app-cache poisoned. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", route="
                + (routeName ?? string.Empty)
                + ", poisonReason="
                + (poisonReason ?? string.Empty)
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

        internal string BuildDiagnosticEventMessage(CaseWorkbookHiddenAppLifecycleDiagnosticEvent diagnosticEvent)
        {
            if (diagnosticEvent == null)
            {
                diagnosticEvent = new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                    "retained-hidden-app-cache-unknown",
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    "diagnostic-event-missing");
            }

            StringBuilder builder = new StringBuilder();
            builder.Append("[KernelFlickerTrace] source=CaseWorkbookOpenStrategy action=")
                .Append(diagnosticEvent.Action)
                .Append(" path=")
                .Append(diagnosticEvent.CaseWorkbookPath);
            AppendField(builder, "route", diagnosticEvent.RouteName);
            AppendField(builder, "caller", diagnosticEvent.Caller);
            AppendField(builder, "eventOutcome", diagnosticEvent.EventOutcome);
            AppendField(builder, "reason", diagnosticEvent.Reason);
            AppendField(builder, "cacheEvent", diagnosticEvent.CacheEvent);
            AppendField(builder, "acquisitionKind", diagnosticEvent.AcquisitionKind);
            AppendField(builder, "returnOutcome", diagnosticEvent.ReturnOutcome);
            AppendField(builder, "poisonReason", diagnosticEvent.PoisonReason);
            AppendField(builder, "cleanupReason", diagnosticEvent.CleanupReason);
            AppendField(builder, "fallbackRoute", diagnosticEvent.FallbackRoute);
            AppendField(builder, "abandonedOperation", diagnosticEvent.AbandonedOperation);
            AppendField(builder, "safetyAction", diagnosticEvent.SafetyAction);
            AppendField(builder, "appHwnd", diagnosticEvent.AppHwnd);
            AppendField(builder, "exceptionType", diagnosticEvent.ExceptionType);

            if (diagnosticEvent.ReusedApplication.HasValue)
            {
                AppendField(builder, "reusedApplication", diagnosticEvent.ReusedApplication.Value.ToString());
            }

            if (diagnosticEvent.RetainedInstancePresent.HasValue)
            {
                AppendField(builder, "retainedInstancePresent", diagnosticEvent.RetainedInstancePresent.Value.ToString());
            }

            if (diagnosticEvent.AppQuitAttempted.HasValue)
            {
                AppendField(builder, "appQuitAttempted", diagnosticEvent.AppQuitAttempted.Value.ToString());
            }

            if (diagnosticEvent.AppQuitCompleted.HasValue)
            {
                AppendField(builder, "appQuitCompleted", diagnosticEvent.AppQuitCompleted.Value.ToString());
            }

            if (diagnosticEvent.ElapsedMilliseconds.HasValue)
            {
                AppendField(
                    builder,
                    "elapsedMs",
                    diagnosticEvent.ElapsedMilliseconds.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (!string.IsNullOrEmpty(diagnosticEvent.ApplicationOwnerFacts))
            {
                builder.Append(",").Append(diagnosticEvent.ApplicationOwnerFacts);
            }

            if (diagnosticEvent.ExpirationDecision != null)
            {
                builder.Append(",").Append(diagnosticEvent.ExpirationDecision.DiagnosticDetails);
            }
            else if (diagnosticEvent.LifecycleFacts != null)
            {
                builder.Append(",").Append(diagnosticEvent.LifecycleFacts.DiagnosticDetails);
            }

            return builder.ToString();
        }

        private static void AppendField(StringBuilder builder, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            builder.Append(",").Append(name).Append("=").Append(value);
        }
    }

    internal sealed class CaseWorkbookHiddenAppLifecycleDiagnosticEvent
    {
        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
            string action,
            string caseWorkbookPath,
            string routeName,
            string caller,
            string reason)
        {
            Action = action ?? string.Empty;
            CaseWorkbookPath = caseWorkbookPath ?? string.Empty;
            RouteName = routeName ?? string.Empty;
            Caller = caller ?? string.Empty;
            Reason = reason ?? string.Empty;
        }

        internal string Action { get; }

        internal string CaseWorkbookPath { get; }

        internal string RouteName { get; }

        internal string Caller { get; }

        internal string Reason { get; }

        internal string EventOutcome { get; set; }

        internal string CacheEvent { get; set; }

        internal string AcquisitionKind { get; set; }

        internal string ReturnOutcome { get; set; }

        internal string PoisonReason { get; set; }

        internal string CleanupReason { get; set; }

        internal string FallbackRoute { get; set; }

        internal string AbandonedOperation { get; set; }

        internal string SafetyAction { get; set; }

        internal string AppHwnd { get; set; }

        internal string ExceptionType { get; set; }

        internal string ApplicationOwnerFacts { get; set; }

        internal bool? ReusedApplication { get; set; }

        internal bool? RetainedInstancePresent { get; set; }

        internal bool? AppQuitAttempted { get; set; }

        internal bool? AppQuitCompleted { get; set; }

        internal long? ElapsedMilliseconds { get; set; }

        internal CaseWorkbookHiddenAppLifecycleFacts LifecycleFacts { get; set; }

        internal CaseWorkbookHiddenAppExpirationDecision ExpirationDecision { get; set; }
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
