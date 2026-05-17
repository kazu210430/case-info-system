using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookHiddenAppCacheDiagnosticEventFactory
    {
        private readonly CaseWorkbookOpenRouteDecisionService _routeDecisionService;

        internal CaseWorkbookHiddenAppCacheDiagnosticEventFactory(CaseWorkbookOpenRouteDecisionService routeDecisionService)
        {
            _routeDecisionService = routeDecisionService ?? throw new ArgumentNullException(nameof(routeDecisionService));
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateShutdownEvent(
            bool retainedInstancePresent,
            string appHwnd)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionShutdown,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.ShutdownHiddenApplicationCache",
                "shutdown-cleanup")
            {
                EventOutcome = retainedInstancePresent ? "dispose-retained-instance" : "no-retained-instance",
                CacheEvent = "shutdown",
                RetainedInstancePresent = retainedInstancePresent,
                AppHwnd = appHwnd,
                CleanupReason = "shutdown-cleanup",
                SafetyAction = retainedInstancePresent ? "dispose-retained-application" : "none",
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateAcquireUnhealthyEvent(
            string caseWorkbookPath,
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts)
        {
            CaseWorkbookHiddenAppLifecycleFacts safeFacts = lifecycleFacts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionOrphanSuspicion,
                caseWorkbookPath,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                "acquire-health-check-failed")
            {
                EventOutcome = "suspected-unhealthy-retained-instance",
                CacheEvent = "acquire-health-check",
                AppHwnd = safeFacts.AppHwnd,
                SafetyAction = "dispose-retained-application",
                AbandonedOperation = "reuse-retained-application",
                LifecycleFacts = safeFacts,
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateAcquireFallbackEvent(
            string caseWorkbookPath,
            CaseWorkbookOpenRouteDecision fallbackDecision,
            long elapsedMilliseconds)
        {
            CaseWorkbookOpenRouteDecision safeDecision = fallbackDecision
                ?? _routeDecisionService.DecideHiddenApplicationCacheAcquisition(cachedApplicationInUse: true);
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionFallback,
                caseWorkbookPath,
                safeDecision.RouteName,
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                safeDecision.Reason)
            {
                EventOutcome = "fallback-to-dedicated-hidden-session",
                CacheEvent = "acquire-fallback",
                FallbackRoute = safeDecision.RouteName,
                AbandonedOperation = "retained-cache-acquire",
                SafetyAction = "open-dedicated-hidden-session",
                ElapsedMilliseconds = elapsedMilliseconds,
                ApplicationOwnerFacts = OwnerFacts(safeDecision.RouteName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateAcquireEvent(
            string caseWorkbookPath,
            string acquisitionReason,
            bool reusedApplication,
            string appHwnd,
            long elapsedMilliseconds)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
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
                AppHwnd = appHwnd,
                ElapsedMilliseconds = elapsedMilliseconds,
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateAcquireOpenFailedPoisonEvent(
            string caseWorkbookPath,
            string appHwnd,
            string exceptionType)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionPoisonMark,
                caseWorkbookPath,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                "workbook-open-failed")
            {
                EventOutcome = "poison-requested",
                CacheEvent = "acquire-open-failed",
                PoisonReason = "workbookOpenFailed",
                AppHwnd = appHwnd,
                ExceptionType = exceptionType,
                SafetyAction = "poison-dispose",
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateReturnSlotMismatchEvent(
            string caseWorkbookPath,
            string routeName,
            string appHwnd,
            bool retainedInstancePresent,
            long? elapsedMilliseconds)
        {
            return CreateIdleReturnEvent(
                caseWorkbookPath,
                routeName,
                "cache-slot-mismatch",
                eventOutcome: "not-returned",
                returnOutcome: "not-cache-slot",
                appHwnd: appHwnd,
                safetyAction: "return-false",
                elapsedMilliseconds: elapsedMilliseconds,
                lifecycleFacts: null)
            .WithRetainedInstancePresent(retainedInstancePresent);
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateReturnFeatureDisabledEvent(
            string caseWorkbookPath,
            string routeName,
            string appHwnd,
            long? elapsedMilliseconds)
        {
            return CreateIdleReturnEvent(
                caseWorkbookPath,
                routeName,
                "feature-flag-disabled",
                eventOutcome: "not-returned",
                returnOutcome: "discard",
                appHwnd: appHwnd,
                safetyAction: "poison-dispose",
                elapsedMilliseconds: elapsedMilliseconds,
                lifecycleFacts: null);
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateReturnHiddenStateFailedEvent(
            string caseWorkbookPath,
            string routeName,
            string appHwnd,
            string exceptionType,
            long? elapsedMilliseconds)
        {
            CaseWorkbookHiddenAppLifecycleDiagnosticEvent diagnosticEvent = CreateIdleReturnEvent(
                caseWorkbookPath,
                routeName,
                "hidden-state-reapply-failed",
                eventOutcome: "not-returned",
                returnOutcome: "discard",
                appHwnd: appHwnd,
                safetyAction: "poison-dispose",
                elapsedMilliseconds: elapsedMilliseconds,
                lifecycleFacts: null);
            diagnosticEvent.ExceptionType = exceptionType;
            return diagnosticEvent;
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateReturnHealthCheckSuspicionEvent(
            string caseWorkbookPath,
            string routeName,
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts)
        {
            CaseWorkbookHiddenAppLifecycleFacts safeFacts = lifecycleFacts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionOrphanSuspicion,
                caseWorkbookPath,
                routeName,
                "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                "return-to-idle-health-check-failed")
            {
                EventOutcome = "suspected-unhealthy-retained-instance",
                CacheEvent = "idle-return-health-check",
                AppHwnd = safeFacts.AppHwnd,
                SafetyAction = "poison-dispose",
                AbandonedOperation = "return-retained-application-to-idle",
                LifecycleFacts = safeFacts,
                ApplicationOwnerFacts = OwnerFacts(routeName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateReturnHealthCheckFailedEvent(
            string caseWorkbookPath,
            string routeName,
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts,
            long? elapsedMilliseconds)
        {
            CaseWorkbookHiddenAppLifecycleFacts safeFacts = lifecycleFacts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            return CreateIdleReturnEvent(
                caseWorkbookPath,
                routeName,
                "health-check-failed",
                eventOutcome: "not-returned",
                returnOutcome: "discard",
                appHwnd: safeFacts.AppHwnd,
                safetyAction: "poison-dispose",
                elapsedMilliseconds: elapsedMilliseconds,
                lifecycleFacts: safeFacts);
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateReturnedToIdleEvent(
            string caseWorkbookPath,
            string routeName,
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts,
            long? elapsedMilliseconds)
        {
            CaseWorkbookHiddenAppLifecycleFacts safeFacts = lifecycleFacts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            return CreateIdleReturnEvent(
                caseWorkbookPath,
                routeName,
                "returnedToIdle",
                eventOutcome: "returned-to-idle",
                returnOutcome: "returned-to-idle",
                appHwnd: safeFacts.AppHwnd,
                safetyAction: "keep-retained-application-idle",
                elapsedMilliseconds: elapsedMilliseconds,
                lifecycleFacts: safeFacts);
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreatePoisonMarkedEvent(
            string caseWorkbookPath,
            string routeName,
            string poisonReason,
            string appHwnd,
            string exceptionType,
            long? elapsedMilliseconds)
        {
            return CreatePoisonEvent(
                caseWorkbookPath,
                routeName,
                poisonReason,
                appHwnd,
                exceptionType,
                eventOutcome: "marked",
                safetyAction: "detach-and-dispose-retained-application",
                elapsedMilliseconds: elapsedMilliseconds);
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreatePoisonSlotMissingEvent(
            string caseWorkbookPath,
            string routeName,
            string poisonReason,
            string appHwnd,
            string exceptionType,
            long? elapsedMilliseconds)
        {
            return CreatePoisonEvent(
                caseWorkbookPath,
                routeName,
                poisonReason,
                appHwnd,
                exceptionType,
                eventOutcome: "not-marked-slot-missing",
                safetyAction: "skip-dispose-slot-not-found",
                elapsedMilliseconds: elapsedMilliseconds);
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateExpirationInUseEvent(
            string cleanupReason,
            CaseWorkbookHiddenAppExpirationDecision expirationDecision)
        {
            return CreateExpirationEvent(
                cleanupReason,
                expirationDecision,
                eventOutcome: "fallback-stop-idle-timer",
                appHwnd: string.Empty,
                abandonedOperation: "idle-timeout-cleanup",
                safetyAction: "stop-idle-timer");
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateExpirationDisposeEvent(
            string cleanupReason,
            string appHwnd,
            CaseWorkbookHiddenAppExpirationDecision expirationDecision)
        {
            return CreateExpirationEvent(
                cleanupReason,
                expirationDecision,
                eventOutcome: "dispose-retained-instance",
                appHwnd: appHwnd,
                abandonedOperation: "retain-idle-cache",
                safetyAction: "dispose-retained-application");
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateDisposeSkippedNotOwnedEvent(
            string cleanupReason,
            string appHwnd,
            bool retainedInstancePresent)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionDispose,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.DisposeCachedHiddenApplicationSlot",
                cleanupReason)
            {
                EventOutcome = "skipped-not-cache-owned",
                CacheEvent = "dispose",
                CleanupReason = cleanupReason,
                AppHwnd = appHwnd,
                RetainedInstancePresent = retainedInstancePresent,
                AppQuitAttempted = false,
                AppQuitCompleted = false,
                SafetyAction = "skip-quit",
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        internal CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateDisposeEvent(
            string cleanupReason,
            string appHwnd,
            bool retainedInstancePresent,
            bool appQuitAttempted,
            bool appQuitCompleted)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionDispose,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.DisposeCachedHiddenApplicationSlot",
                cleanupReason)
            {
                EventOutcome = appQuitCompleted ? "disposed" : "degraded",
                CacheEvent = "dispose",
                CleanupReason = cleanupReason,
                AppHwnd = appHwnd,
                RetainedInstancePresent = retainedInstancePresent,
                AppQuitAttempted = appQuitAttempted,
                AppQuitCompleted = appQuitCompleted,
                SafetyAction = appQuitCompleted ? "quit-and-release" : "release-after-quit-failure",
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        private CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateIdleReturnEvent(
            string caseWorkbookPath,
            string routeName,
            string reason,
            string eventOutcome,
            string returnOutcome,
            string appHwnd,
            string safetyAction,
            long? elapsedMilliseconds,
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionIdleReturn,
                caseWorkbookPath,
                routeName,
                "CaseWorkbookOpenStrategy.TryReturnCachedHiddenApplicationToIdle",
                reason)
            {
                EventOutcome = eventOutcome,
                CacheEvent = "idle-return",
                ReturnOutcome = returnOutcome,
                AppHwnd = appHwnd,
                SafetyAction = safetyAction,
                LifecycleFacts = lifecycleFacts,
                ElapsedMilliseconds = elapsedMilliseconds,
                ApplicationOwnerFacts = OwnerFacts(routeName)
            };
        }

        private CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreatePoisonEvent(
            string caseWorkbookPath,
            string routeName,
            string poisonReason,
            string appHwnd,
            string exceptionType,
            string eventOutcome,
            string safetyAction,
            long? elapsedMilliseconds)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionPoisonMark,
                caseWorkbookPath,
                routeName,
                "CaseWorkbookOpenStrategy.MarkCachedHiddenApplicationPoisoned",
                poisonReason)
            {
                EventOutcome = eventOutcome,
                CacheEvent = "poison",
                PoisonReason = poisonReason,
                AppHwnd = appHwnd,
                ExceptionType = exceptionType,
                SafetyAction = safetyAction,
                ElapsedMilliseconds = elapsedMilliseconds,
                ApplicationOwnerFacts = OwnerFacts(routeName)
            };
        }

        private CaseWorkbookHiddenAppLifecycleDiagnosticEvent CreateExpirationEvent(
            string cleanupReason,
            CaseWorkbookHiddenAppExpirationDecision expirationDecision,
            string eventOutcome,
            string appHwnd,
            string abandonedOperation,
            string safetyAction)
        {
            return new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionTimeoutFallback,
                string.Empty,
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                "CaseWorkbookOpenStrategy.CleanupExpiredCachedHiddenApplicationUnlocked",
                expirationDecision == null ? string.Empty : expirationDecision.DecisionReason)
            {
                EventOutcome = eventOutcome,
                CacheEvent = "expiration-decision",
                CleanupReason = cleanupReason,
                AppHwnd = appHwnd,
                AbandonedOperation = abandonedOperation,
                SafetyAction = safetyAction,
                ExpirationDecision = expirationDecision,
                ApplicationOwnerFacts = OwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName)
            };
        }

        private string OwnerFacts(string routeName)
        {
            return _routeDecisionService.BuildApplicationOwnerFacts(routeName);
        }
    }

    internal static class CaseWorkbookHiddenAppLifecycleDiagnosticEventExtensions
    {
        internal static CaseWorkbookHiddenAppLifecycleDiagnosticEvent WithRetainedInstancePresent(
            this CaseWorkbookHiddenAppLifecycleDiagnosticEvent diagnosticEvent,
            bool retainedInstancePresent)
        {
            if (diagnosticEvent != null)
            {
                diagnosticEvent.RetainedInstancePresent = retainedInstancePresent;
            }

            return diagnosticEvent;
        }
    }
}
