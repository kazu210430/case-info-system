using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookOpenRouteDecisionService
    {
        internal const string DedicatedHiddenInnerSaveEnvironmentVariableName = "CASEINFO_EXPERIMENT_DEDICATED_HIDDEN_INNER_SAVE";
        internal const string LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName = "CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL";
        internal const string HiddenApplicationCacheEnvironmentVariableName = "CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE";
        internal const string HiddenApplicationCacheIdleSecondsEnvironmentVariableName = "CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE_IDLE_SECONDS";
        internal const string LegacyHiddenRouteName = "legacy-isolated";
        internal const string ExperimentalIsolatedInnerSaveRouteName = "experimental-isolated-inner-save";
        internal const string CreatedCaseDisplayHiddenRouteName = "created-case-display";
        internal const string HiddenApplicationCacheRouteName = "app-cache";
        internal const string HiddenApplicationCacheBypassInUseRouteName = "app-cache-bypass-inuse";
        internal const int DefaultHiddenApplicationCacheIdleSeconds = 15;
        internal const string ApplicationKindIsolated = "isolated";
        internal const string ApplicationKindRetainedHiddenAppCache = "retained-hidden-app-cache";
        internal const string ApplicationKindSharedCurrent = "shared-current";
        internal const string ApplicationLifetimeOwnerCaseWorkbookOpenStrategy = "CaseWorkbookOpenStrategy";
        internal const string ApplicationLifetimeOwnerUserOrExcelHost = "user-or-excel-host";

        internal CaseWorkbookOpenRouteDecision DecideHiddenCreateRoute()
        {
            return DecideHiddenCreateRoute(ReadHiddenRouteSwitches());
        }

        internal CaseWorkbookOpenRouteDecision DecideHiddenCreateRoute(CaseWorkbookOpenRouteSwitches switches)
        {
            CaseWorkbookOpenRouteSwitches effectiveSwitches = switches ?? ReadHiddenRouteSwitches();
            if (effectiveSwitches.HiddenApplicationCacheEnabled)
            {
                return CreateRouteDecision(
                    HiddenApplicationCacheRouteName,
                    "hiddenApplicationCacheEnabled",
                    saveBeforeClose: false,
                    useHiddenApplicationCache: true,
                    isFallbackRoute: false);
            }

            if (effectiveSwitches.DedicatedHiddenInnerSaveEnabled)
            {
                return CreateRouteDecision(
                    ExperimentalIsolatedInnerSaveRouteName,
                    effectiveSwitches.LegacyDedicatedHiddenInnerSaveAliasEnabled
                        ? "legacyDedicatedHiddenInnerSaveAliasEnabled"
                        : "dedicatedHiddenInnerSaveEnabled",
                    saveBeforeClose: true,
                    useHiddenApplicationCache: false,
                    isFallbackRoute: false);
            }

            return CreateRouteDecision(
                LegacyHiddenRouteName,
                "defaultLegacyHiddenRoute",
                saveBeforeClose: false,
                useHiddenApplicationCache: false,
                isFallbackRoute: false);
        }

        internal CaseWorkbookOpenRouteDecision DecideHiddenApplicationCacheAcquisition(bool cachedApplicationInUse)
        {
            return cachedApplicationInUse
                ? CreateRouteDecision(
                    HiddenApplicationCacheBypassInUseRouteName,
                    "hiddenApplicationCacheInUse",
                    saveBeforeClose: false,
                    useHiddenApplicationCache: false,
                    isFallbackRoute: true)
                : CreateRouteDecision(
                    HiddenApplicationCacheRouteName,
                    "hiddenApplicationCacheAvailable",
                    saveBeforeClose: false,
                    useHiddenApplicationCache: true,
                    isFallbackRoute: false);
        }

        internal CaseWorkbookOpenRouteDecision DecideCreatedCaseDisplayRoute()
        {
            return CreateRouteDecision(
                CreatedCaseDisplayHiddenRouteName,
                "createdCaseDisplayHandoff",
                saveBeforeClose: false,
                useHiddenApplicationCache: false,
                isFallbackRoute: false);
        }

        internal CaseWorkbookOpenRouteSwitches ReadHiddenRouteSwitches()
        {
            string dedicatedHiddenInnerSaveValue = Environment.GetEnvironmentVariable(DedicatedHiddenInnerSaveEnvironmentVariableName);
            string legacyDedicatedHiddenInnerSaveAliasValue = Environment.GetEnvironmentVariable(LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName);
            string hiddenApplicationCacheValue = Environment.GetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName);
            string hiddenApplicationCacheIdleSecondsValue = Environment.GetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName);
            bool dedicatedHiddenInnerSaveEnabled = IsEnabledSwitchValue(dedicatedHiddenInnerSaveValue);
            bool legacyDedicatedHiddenInnerSaveAliasEnabled = IsEnabledSwitchValue(legacyDedicatedHiddenInnerSaveAliasValue);
            bool hiddenApplicationCacheEnabled = IsEnabledSwitchValue(hiddenApplicationCacheValue);
            int idleSeconds = ResolveHiddenApplicationCacheIdleSeconds(hiddenApplicationCacheIdleSecondsValue);

            return new CaseWorkbookOpenRouteSwitches(
                dedicatedHiddenInnerSaveEnabled || legacyDedicatedHiddenInnerSaveAliasEnabled,
                legacyDedicatedHiddenInnerSaveAliasEnabled,
                hiddenApplicationCacheEnabled,
                idleSeconds,
                dedicatedHiddenInnerSaveValue,
                legacyDedicatedHiddenInnerSaveAliasValue,
                hiddenApplicationCacheValue,
                hiddenApplicationCacheIdleSecondsValue);
        }

        internal bool IsHiddenApplicationCacheEnabled()
        {
            return ReadHiddenRouteSwitches().HiddenApplicationCacheEnabled;
        }

        internal int ResolveHiddenApplicationCacheIdleSeconds()
        {
            return ReadHiddenRouteSwitches().HiddenApplicationCacheIdleSeconds;
        }

        internal string BuildApplicationOwnerFacts(string routeName)
        {
            bool isSharedCurrentApp = IsSharedCurrentApplicationRoute(routeName);
            bool isRetainedHiddenAppCache = IsRetainedHiddenApplicationCacheRoute(routeName);
            bool isIsolatedApp = !isSharedCurrentApp && !isRetainedHiddenAppCache;
            return "applicationKind=" + ResolveApplicationKind(routeName)
                + ",applicationLifetimeOwner=" + ResolveApplicationLifetimeOwner(routeName)
                + ",isSharedCurrentApp=" + isSharedCurrentApp.ToString()
                + ",isIsolatedApp=" + isIsolatedApp.ToString()
                + ",isRetainedHiddenAppCache=" + isRetainedHiddenAppCache.ToString();
        }

        private static CaseWorkbookOpenRouteDecision CreateRouteDecision(
            string routeName,
            string reason,
            bool saveBeforeClose,
            bool useHiddenApplicationCache,
            bool isFallbackRoute)
        {
            bool isSharedCurrentApp = IsSharedCurrentApplicationRoute(routeName);
            bool isRetainedHiddenAppCache = IsRetainedHiddenApplicationCacheRoute(routeName);
            bool isIsolatedApp = !isSharedCurrentApp && !isRetainedHiddenAppCache;
            string applicationKind = ResolveApplicationKind(routeName);
            string applicationLifetimeOwner = ResolveApplicationLifetimeOwner(routeName);
            string applicationOwnerFacts = "applicationKind=" + applicationKind
                + ",applicationLifetimeOwner=" + applicationLifetimeOwner
                + ",isSharedCurrentApp=" + isSharedCurrentApp.ToString()
                + ",isIsolatedApp=" + isIsolatedApp.ToString()
                + ",isRetainedHiddenAppCache=" + isRetainedHiddenAppCache.ToString();
            string routeTraceDetails = "route=" + (routeName ?? string.Empty)
                + ",routeReason=" + (reason ?? string.Empty)
                + ",saveBeforeClose=" + saveBeforeClose.ToString()
                + ",useHiddenApplicationCache=" + useHiddenApplicationCache.ToString()
                + ",isFallbackRoute=" + isFallbackRoute.ToString()
                + "," + applicationOwnerFacts;

            return new CaseWorkbookOpenRouteDecision(
                routeName,
                reason,
                routeTraceDetails,
                applicationOwnerFacts,
                applicationKind,
                applicationLifetimeOwner,
                isSharedCurrentApp,
                isIsolatedApp,
                isRetainedHiddenAppCache,
                saveBeforeClose,
                useHiddenApplicationCache,
                isFallbackRoute);
        }

        private static bool IsEnabledSwitchValue(string value)
        {
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static int ResolveHiddenApplicationCacheIdleSeconds(string value)
        {
            int parsed;
            if (int.TryParse(value, out parsed) && parsed > 0)
            {
                return parsed;
            }

            return DefaultHiddenApplicationCacheIdleSeconds;
        }

        private static string ResolveApplicationKind(string routeName)
        {
            if (IsSharedCurrentApplicationRoute(routeName))
            {
                return ApplicationKindSharedCurrent;
            }

            if (IsRetainedHiddenApplicationCacheRoute(routeName))
            {
                return ApplicationKindRetainedHiddenAppCache;
            }

            return ApplicationKindIsolated;
        }

        private static string ResolveApplicationLifetimeOwner(string routeName)
        {
            return IsSharedCurrentApplicationRoute(routeName)
                ? ApplicationLifetimeOwnerUserOrExcelHost
                : ApplicationLifetimeOwnerCaseWorkbookOpenStrategy;
        }

        private static bool IsSharedCurrentApplicationRoute(string routeName)
        {
            return string.Equals(routeName, CreatedCaseDisplayHiddenRouteName, StringComparison.Ordinal);
        }

        private static bool IsRetainedHiddenApplicationCacheRoute(string routeName)
        {
            return string.Equals(routeName, HiddenApplicationCacheRouteName, StringComparison.Ordinal);
        }
    }

    internal sealed class CaseWorkbookOpenRouteSwitches
    {
        internal CaseWorkbookOpenRouteSwitches(
            bool dedicatedHiddenInnerSaveEnabled,
            bool legacyDedicatedHiddenInnerSaveAliasEnabled,
            bool hiddenApplicationCacheEnabled,
            int hiddenApplicationCacheIdleSeconds,
            string dedicatedHiddenInnerSaveRawValue,
            string legacyDedicatedHiddenInnerSaveAliasRawValue,
            string hiddenApplicationCacheRawValue,
            string hiddenApplicationCacheIdleSecondsRawValue)
        {
            DedicatedHiddenInnerSaveEnabled = dedicatedHiddenInnerSaveEnabled;
            LegacyDedicatedHiddenInnerSaveAliasEnabled = legacyDedicatedHiddenInnerSaveAliasEnabled;
            HiddenApplicationCacheEnabled = hiddenApplicationCacheEnabled;
            HiddenApplicationCacheIdleSeconds = hiddenApplicationCacheIdleSeconds;
            DedicatedHiddenInnerSaveRawValue = dedicatedHiddenInnerSaveRawValue ?? string.Empty;
            LegacyDedicatedHiddenInnerSaveAliasRawValue = legacyDedicatedHiddenInnerSaveAliasRawValue ?? string.Empty;
            HiddenApplicationCacheRawValue = hiddenApplicationCacheRawValue ?? string.Empty;
            HiddenApplicationCacheIdleSecondsRawValue = hiddenApplicationCacheIdleSecondsRawValue ?? string.Empty;
        }

        internal bool DedicatedHiddenInnerSaveEnabled { get; }

        internal bool LegacyDedicatedHiddenInnerSaveAliasEnabled { get; }

        internal bool HiddenApplicationCacheEnabled { get; }

        internal int HiddenApplicationCacheIdleSeconds { get; }

        internal string DedicatedHiddenInnerSaveRawValue { get; }

        internal string LegacyDedicatedHiddenInnerSaveAliasRawValue { get; }

        internal string HiddenApplicationCacheRawValue { get; }

        internal string HiddenApplicationCacheIdleSecondsRawValue { get; }
    }

    internal sealed class CaseWorkbookOpenRouteDecision
    {
        internal CaseWorkbookOpenRouteDecision(
            string routeName,
            string reason,
            string routeTraceDetails,
            string applicationOwnerFacts,
            string applicationKind,
            string applicationLifetimeOwner,
            bool isSharedCurrentApplication,
            bool isIsolatedApplication,
            bool isRetainedHiddenApplicationCache,
            bool saveBeforeClose,
            bool useHiddenApplicationCache,
            bool isFallbackRoute)
        {
            RouteName = routeName ?? string.Empty;
            Reason = reason ?? string.Empty;
            RouteTraceDetails = routeTraceDetails ?? string.Empty;
            ApplicationOwnerFacts = applicationOwnerFacts ?? string.Empty;
            ApplicationKind = applicationKind ?? string.Empty;
            ApplicationLifetimeOwner = applicationLifetimeOwner ?? string.Empty;
            IsSharedCurrentApplication = isSharedCurrentApplication;
            IsIsolatedApplication = isIsolatedApplication;
            IsRetainedHiddenApplicationCache = isRetainedHiddenApplicationCache;
            SaveBeforeClose = saveBeforeClose;
            UseHiddenApplicationCache = useHiddenApplicationCache;
            IsFallbackRoute = isFallbackRoute;
        }

        internal string RouteName { get; }

        internal string Reason { get; }

        internal string RouteTraceDetails { get; }

        internal string ApplicationOwnerFacts { get; }

        internal string ApplicationKind { get; }

        internal string ApplicationLifetimeOwner { get; }

        internal bool IsSharedCurrentApplication { get; }

        internal bool IsIsolatedApplication { get; }

        internal bool IsRetainedHiddenApplicationCache { get; }

        internal bool SaveBeforeClose { get; }

        internal bool UseHiddenApplicationCache { get; }

        internal bool IsFallbackRoute { get; }
    }
}
