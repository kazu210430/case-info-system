using System.Globalization;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class ExcelWindowRestoreDecision
    {
        internal ExcelWindowRestoreDecision(bool shouldRestore, bool restoreSkipped, string restoreSkipReason)
        {
            ShouldRestore = shouldRestore;
            RestoreSkipped = ExcelWindowRestoreDecisionPolicy.FormatBooleanLike(restoreSkipped);
            RestoreSkipReason = restoreSkipReason ?? string.Empty;
        }

        internal bool ShouldRestore { get; }

        internal string RestoreSkipped { get; }

        internal string RestoreSkipReason { get; }
    }

    internal static class ExcelWindowRestoreDecisionPolicy
    {
        internal const int SwHide = 0;
        internal const int SwShowNormal = 1;
        internal const int SwShowMinimized = 2;
        internal const int SwShowMaximized = 3;
        internal const int SwShowNoActivate = 4;
        internal const int SwShow = 5;
        internal const int SwMinimize = 6;
        internal const int SwShowMinNoActive = 7;
        internal const int SwShowNa = 8;
        internal const int SwRestore = 9;
        internal const int SwShowDefault = 10;
        internal const int SwForceMinimize = 11;

        internal static ExcelWindowRestoreDecision Decide(
            bool windowResolved,
            bool visibleReadSucceeded,
            bool visible,
            bool placementReadSucceeded,
            int showCmd)
        {
            if (!windowResolved)
            {
                return RestoreRequired("restore-required:window-null");
            }

            if (!visibleReadSucceeded)
            {
                return RestoreRequired("restore-required:visible-read-failed");
            }

            if (!visible)
            {
                return RestoreRequired("restore-required:not-visible");
            }

            if (!placementReadSucceeded)
            {
                return RestoreRequired("restore-required:placement-read-failed");
            }

            if (showCmd == SwHide)
            {
                return RestoreRequired("restore-required:hidden");
            }

            if (IsPlacementMinimized(showCmd))
            {
                return RestoreRequired("restore-required:minimized");
            }

            if (IsPlacementMaximized(showCmd))
            {
                return RestoreRequired("restore-required:maximized");
            }

            if (showCmd == SwShowNormal)
            {
                return new ExcelWindowRestoreDecision(
                    shouldRestore: false,
                    restoreSkipped: true,
                    restoreSkipReason: "visible-shownormal-not-minimized-not-maximized");
            }

            return RestoreRequired("restore-required:showCmd=" + ResolveShowCmdName(showCmd));
        }

        internal static string FormatShowCmd(int showCmd)
        {
            return ResolveShowCmdName(showCmd)
                + "("
                + showCmd.ToString(CultureInfo.InvariantCulture)
                + ")";
        }

        internal static string ResolveShowCmdName(int showCmd)
        {
            switch (showCmd)
            {
                case SwHide:
                    return "SW_HIDE";
                case SwShowNormal:
                    return "SW_SHOWNORMAL";
                case SwShowMinimized:
                    return "SW_SHOWMINIMIZED";
                case SwShowMaximized:
                    return "SW_SHOWMAXIMIZED";
                case SwShowNoActivate:
                    return "SW_SHOWNOACTIVATE";
                case SwShow:
                    return "SW_SHOW";
                case SwMinimize:
                    return "SW_MINIMIZE";
                case SwShowMinNoActive:
                    return "SW_SHOWMINNOACTIVE";
                case SwShowNa:
                    return "SW_SHOWNA";
                case SwRestore:
                    return "SW_RESTORE";
                case SwShowDefault:
                    return "SW_SHOWDEFAULT";
                case SwForceMinimize:
                    return "SW_FORCEMINIMIZE";
                default:
                    return "SW_UNKNOWN";
            }
        }

        internal static bool IsPlacementMinimized(int showCmd)
        {
            return showCmd == SwShowMinimized
                || showCmd == SwMinimize
                || showCmd == SwShowMinNoActive
                || showCmd == SwForceMinimize;
        }

        internal static bool IsPlacementMaximized(int showCmd)
        {
            return showCmd == SwShowMaximized;
        }

        internal static bool IsPlacementNormal(int showCmd)
        {
            return !IsPlacementMinimized(showCmd) && !IsPlacementMaximized(showCmd);
        }

        internal static string FormatBooleanLike(bool value)
        {
            return value ? "True" : "False";
        }

        private static ExcelWindowRestoreDecision RestoreRequired(string restoreSkipReason)
        {
            return new ExcelWindowRestoreDecision(
                shouldRestore: true,
                restoreSkipped: false,
                restoreSkipReason: restoreSkipReason);
        }
    }
}
