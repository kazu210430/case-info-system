using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookOpenCleanupOutcomeService
    {
        internal const string HiddenExcelCleanupCompleted = "HiddenExcelCleanupCompleted";
        internal const string HiddenExcelCleanupNotRequired = "HiddenExcelCleanupNotRequired";
        internal const string HiddenExcelCleanupDegraded = "HiddenExcelCleanupDegraded";
        internal const string HiddenExcelCleanupFailed = "HiddenExcelCleanupFailed";
        internal const string IsolatedAppReleased = "IsolatedAppReleased";
        internal const string IsolatedAppReleaseNotRequired = "IsolatedAppReleaseNotRequired";
        internal const string IsolatedAppReleaseDegraded = "IsolatedAppReleaseDegraded";
        internal const string IsolatedAppReleaseFailed = "IsolatedAppReleaseFailed";
        internal const string RetainedInstanceReturnedToIdle = "RetainedInstanceReturnedToIdle";
        internal const string RetainedInstancePoisoned = "RetainedInstancePoisoned";
        internal const string RetainedInstanceCleanupCompleted = "RetainedInstanceCleanupCompleted";
        internal const string RetainedInstanceCleanupSkipped = "RetainedInstanceCleanupSkipped";
        internal const string RetainedInstanceCleanupDegraded = "RetainedInstanceCleanupDegraded";
        internal const string RetainedInstanceCleanupNotRequired = "RetainedInstanceCleanupNotRequired";

        private readonly CaseWorkbookOpenRouteDecisionService _routeDecisionService;

        internal CaseWorkbookOpenCleanupOutcomeService(CaseWorkbookOpenRouteDecisionService routeDecisionService)
        {
            _routeDecisionService = routeDecisionService ?? throw new ArgumentNullException(nameof(routeDecisionService));
        }

        internal CaseWorkbookOpenHiddenCleanupOutcome CreateDedicatedHiddenSessionOutcome(
            string owner,
            string caseWorkbookPath,
            string routeName,
            CaseWorkbookOpenCleanupFacts facts,
            bool cleanupFailed)
        {
            CaseWorkbookOpenCleanupFacts safeFacts = facts ?? CaseWorkbookOpenCleanupFacts.Empty;
            return CreateHiddenCleanupOutcome(
                owner,
                caseWorkbookPath,
                routeName,
                ResolveDedicatedHiddenCleanupOutcome(safeFacts, cleanupFailed),
                ResolveIsolatedAppReleaseOutcome(safeFacts),
                string.Empty,
                safeFacts,
                cacheReturnedToIdle: false,
                cachePoisoned: false,
                cleanupFailed ? "cleanupException" : "dedicatedSessionFinalized");
        }

        internal CaseWorkbookOpenHiddenCleanupOutcome CreateCachedHiddenSessionReturnedToIdleOutcome(
            string owner,
            string caseWorkbookPath,
            string routeName,
            CaseWorkbookOpenCleanupFacts facts)
        {
            return CreateHiddenCleanupOutcome(
                owner,
                caseWorkbookPath,
                routeName,
                HiddenExcelCleanupCompleted,
                string.Empty,
                RetainedInstanceReturnedToIdle,
                facts ?? CaseWorkbookOpenCleanupFacts.Empty,
                cacheReturnedToIdle: true,
                cachePoisoned: false,
                "returnedToIdle");
        }

        internal CaseWorkbookOpenHiddenCleanupOutcome CreateCachedHiddenSessionPoisonedOutcome(
            string owner,
            string caseWorkbookPath,
            string routeName,
            CaseWorkbookOpenCleanupFacts facts,
            string reason)
        {
            CaseWorkbookOpenCleanupFacts safeFacts = facts ?? CaseWorkbookOpenCleanupFacts.Empty;
            return CreateHiddenCleanupOutcome(
                owner,
                caseWorkbookPath,
                routeName,
                ResolveCachedHiddenCleanupOutcome(safeFacts),
                string.Empty,
                RetainedInstancePoisoned,
                safeFacts,
                cacheReturnedToIdle: false,
                cachePoisoned: true,
                reason);
        }

        internal CaseWorkbookOpenRetainedCleanupOutcome CreateRetainedInstanceCleanupOutcome(
            string reason,
            string appHwnd,
            bool retainedInstancePresent,
            bool isOwnedByCache,
            bool quitAttempted,
            bool quitCompleted)
        {
            string retainedInstanceOutcome = ResolveRetainedInstanceCleanupOutcome(
                retainedInstancePresent,
                isOwnedByCache,
                quitCompleted);
            return new CaseWorkbookOpenRetainedCleanupOutcome(
                _routeDecisionService.BuildApplicationOwnerFacts(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName),
                retainedInstanceOutcome,
                reason,
                appHwnd,
                quitAttempted,
                quitCompleted);
        }

        private CaseWorkbookOpenHiddenCleanupOutcome CreateHiddenCleanupOutcome(
            string owner,
            string caseWorkbookPath,
            string routeName,
            string hiddenCleanupOutcome,
            string isolatedAppOutcome,
            string retainedInstanceOutcome,
            CaseWorkbookOpenCleanupFacts facts,
            bool cacheReturnedToIdle,
            bool cachePoisoned,
            string reason)
        {
            string applicationOwnerFacts = _routeDecisionService.BuildApplicationOwnerFacts(routeName);
            string details = "scope=hidden-cleanup"
                + ",route=" + (routeName ?? string.Empty)
                + "," + applicationOwnerFacts
                + ",hiddenCleanupOutcome=" + (hiddenCleanupOutcome ?? string.Empty)
                + ",isolatedAppOutcome=" + (isolatedAppOutcome ?? string.Empty)
                + ",retainedInstanceOutcome=" + (retainedInstanceOutcome ?? string.Empty)
                + ",workbookPresent=" + facts.WorkbookPresent.ToString()
                + ",workbookCloseAttempted=" + facts.WorkbookCloseAttempted.ToString()
                + ",workbookCloseCompleted=" + facts.WorkbookCloseCompleted.ToString()
                + ",appPresent=" + facts.AppPresent.ToString()
                + ",appQuitAttempted=" + facts.AppQuitAttempted.ToString()
                + ",appQuitCompleted=" + facts.AppQuitCompleted.ToString()
                + ",cacheReturnedToIdle=" + cacheReturnedToIdle.ToString()
                + ",cachePoisoned=" + cachePoisoned.ToString()
                + ",outcomeReason=" + (reason ?? string.Empty);

            return new CaseWorkbookOpenHiddenCleanupOutcome(
                owner,
                caseWorkbookPath,
                routeName,
                hiddenCleanupOutcome,
                isolatedAppOutcome,
                retainedInstanceOutcome,
                facts,
                cacheReturnedToIdle,
                cachePoisoned,
                reason,
                details);
        }

        private static string ResolveDedicatedHiddenCleanupOutcome(CaseWorkbookOpenCleanupFacts facts, bool cleanupFailed)
        {
            if (!facts.WorkbookPresent && !facts.AppPresent)
            {
                return HiddenExcelCleanupNotRequired;
            }

            if (cleanupFailed || (facts.WorkbookPresent && !facts.WorkbookCloseCompleted))
            {
                return HiddenExcelCleanupFailed;
            }

            if (facts.AppPresent && !facts.AppQuitCompleted)
            {
                return HiddenExcelCleanupDegraded;
            }

            return HiddenExcelCleanupCompleted;
        }

        private static string ResolveCachedHiddenCleanupOutcome(CaseWorkbookOpenCleanupFacts facts)
        {
            if (!facts.WorkbookPresent)
            {
                return HiddenExcelCleanupNotRequired;
            }

            return facts.WorkbookCloseCompleted
                ? HiddenExcelCleanupCompleted
                : HiddenExcelCleanupDegraded;
        }

        private static string ResolveIsolatedAppReleaseOutcome(CaseWorkbookOpenCleanupFacts facts)
        {
            if (!facts.AppPresent)
            {
                return IsolatedAppReleaseNotRequired;
            }

            if (facts.AppQuitCompleted)
            {
                return IsolatedAppReleased;
            }

            return facts.AppQuitAttempted
                ? IsolatedAppReleaseFailed
                : IsolatedAppReleaseDegraded;
        }

        private static string ResolveRetainedInstanceCleanupOutcome(
            bool retainedInstancePresent,
            bool isOwnedByCache,
            bool quitCompleted)
        {
            if (!retainedInstancePresent)
            {
                return RetainedInstanceCleanupNotRequired;
            }

            if (!isOwnedByCache)
            {
                return RetainedInstanceCleanupSkipped;
            }

            return quitCompleted
                ? RetainedInstanceCleanupCompleted
                : RetainedInstanceCleanupDegraded;
        }
    }

    internal sealed class CaseWorkbookOpenCleanupFacts
    {
        internal static readonly CaseWorkbookOpenCleanupFacts Empty = new CaseWorkbookOpenCleanupFacts(
            workbookPresent: false,
            workbookCloseAttempted: false,
            workbookCloseCompleted: true,
            appPresent: false,
            appQuitAttempted: false,
            appQuitCompleted: true);

        internal CaseWorkbookOpenCleanupFacts(
            bool workbookPresent,
            bool workbookCloseAttempted,
            bool workbookCloseCompleted,
            bool appPresent,
            bool appQuitAttempted,
            bool appQuitCompleted)
        {
            WorkbookPresent = workbookPresent;
            WorkbookCloseAttempted = workbookCloseAttempted;
            WorkbookCloseCompleted = workbookCloseCompleted;
            AppPresent = appPresent;
            AppQuitAttempted = appQuitAttempted;
            AppQuitCompleted = appQuitCompleted;
        }

        internal bool WorkbookPresent { get; }

        internal bool WorkbookCloseAttempted { get; }

        internal bool WorkbookCloseCompleted { get; }

        internal bool AppPresent { get; }

        internal bool AppQuitAttempted { get; }

        internal bool AppQuitCompleted { get; }
    }

    internal sealed class CaseWorkbookOpenHiddenCleanupOutcome
    {
        internal CaseWorkbookOpenHiddenCleanupOutcome(
            string owner,
            string caseWorkbookPath,
            string routeName,
            string hiddenCleanupOutcome,
            string isolatedAppOutcome,
            string retainedInstanceOutcome,
            CaseWorkbookOpenCleanupFacts facts,
            bool cacheReturnedToIdle,
            bool cachePoisoned,
            string reason,
            string details)
        {
            Owner = owner ?? string.Empty;
            CaseWorkbookPath = caseWorkbookPath ?? string.Empty;
            RouteName = routeName ?? string.Empty;
            HiddenCleanupOutcome = hiddenCleanupOutcome ?? string.Empty;
            IsolatedAppOutcome = isolatedAppOutcome ?? string.Empty;
            RetainedInstanceOutcome = retainedInstanceOutcome ?? string.Empty;
            Facts = facts ?? CaseWorkbookOpenCleanupFacts.Empty;
            CacheReturnedToIdle = cacheReturnedToIdle;
            CachePoisoned = cachePoisoned;
            Reason = reason ?? string.Empty;
            Details = details ?? string.Empty;
        }

        internal string Owner { get; }

        internal string CaseWorkbookPath { get; }

        internal string RouteName { get; }

        internal string HiddenCleanupOutcome { get; }

        internal string IsolatedAppOutcome { get; }

        internal string RetainedInstanceOutcome { get; }

        internal CaseWorkbookOpenCleanupFacts Facts { get; }

        internal bool CacheReturnedToIdle { get; }

        internal bool CachePoisoned { get; }

        internal string Reason { get; }

        internal string Details { get; }

        internal string KernelFlickerTraceMessage
        {
            get
            {
                return "[KernelFlickerTrace] source="
                    + Owner
                    + " action=hidden-excel-cleanup-outcome path="
                    + CaseWorkbookPath
                    + ", "
                    + Details;
            }
        }
    }

    internal sealed class CaseWorkbookOpenRetainedCleanupOutcome
    {
        internal CaseWorkbookOpenRetainedCleanupOutcome(
            string applicationOwnerFacts,
            string retainedInstanceOutcome,
            string cleanupReason,
            string appHwnd,
            bool appQuitAttempted,
            bool appQuitCompleted)
        {
            ApplicationOwnerFacts = applicationOwnerFacts ?? string.Empty;
            RetainedInstanceOutcome = retainedInstanceOutcome ?? string.Empty;
            CleanupReason = cleanupReason ?? string.Empty;
            AppHwnd = appHwnd ?? string.Empty;
            AppQuitAttempted = appQuitAttempted;
            AppQuitCompleted = appQuitCompleted;
        }

        internal string ApplicationOwnerFacts { get; }

        internal string RetainedInstanceOutcome { get; }

        internal string CleanupReason { get; }

        internal string AppHwnd { get; }

        internal bool AppQuitAttempted { get; }

        internal bool AppQuitCompleted { get; }

        internal string KernelFlickerTraceMessage
        {
            get
            {
                return "[KernelFlickerTrace] source=CaseWorkbookOpenStrategy.DisposeCachedHiddenApplicationSlot"
                    + " action=retained-instance-cleanup-outcome"
                    + " " + ApplicationOwnerFacts
                    + ", retainedInstanceOutcome=" + RetainedInstanceOutcome
                    + ", cleanupReason=" + CleanupReason
                    + ", appHwnd=" + AppHwnd
                    + ", appQuitAttempted=" + AppQuitAttempted.ToString()
                    + ", appQuitCompleted=" + AppQuitCompleted.ToString();
            }
        }
    }
}
