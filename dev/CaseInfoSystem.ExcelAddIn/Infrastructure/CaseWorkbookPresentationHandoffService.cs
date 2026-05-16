using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookPresentationHandoffService
    {
        internal CaseWorkbookPresentationHandoffPlan CreateHiddenForDisplayPlan(
            string caseWorkbookPath,
            CaseWorkbookOpenRouteDecision routeDecision,
            Excel.Window previousActiveWindow,
            bool previousApplicationVisible,
            bool previousScreenUpdating,
            bool previousEnableEvents,
            bool previousDisplayAlerts)
        {
            if (routeDecision == null)
            {
                throw new ArgumentNullException(nameof(routeDecision));
            }

            var sharedStateFacts = new CaseWorkbookSharedDisplayStateFacts(
                previousActiveWindow,
                previousApplicationVisible,
                previousScreenUpdating,
                previousEnableEvents,
                previousDisplayAlerts);
            CaseWorkbookPreviousWindowRestoreDecision previousWindowRestoreDecision =
                DecidePreviousWindowRestore(sharedStateFacts);
            string diagnosticDetails = "scope=presentation-handoff"
                + ",route=" + routeDecision.RouteName
                + ",routeReason=" + routeDecision.Reason
                + "," + routeDecision.ApplicationOwnerFacts
                + ",previousApplicationVisible=" + previousApplicationVisible.ToString()
                + ",previousWindowRestoreRequired=" + previousWindowRestoreDecision.ShouldRestore.ToString()
                + ",previousWindowRestoreReason=" + previousWindowRestoreDecision.Reason;

            return new CaseWorkbookPresentationHandoffPlan(
                caseWorkbookPath,
                routeDecision,
                sharedStateFacts,
                previousWindowRestoreDecision,
                diagnosticDetails);
        }

        internal CaseWorkbookPreviousWindowRestoreDecision DecidePreviousWindowRestore(
            CaseWorkbookSharedDisplayStateFacts sharedStateFacts)
        {
            CaseWorkbookSharedDisplayStateFacts safeFacts = sharedStateFacts ?? CaseWorkbookSharedDisplayStateFacts.Empty;
            if (!safeFacts.PreviousApplicationVisible)
            {
                return new CaseWorkbookPreviousWindowRestoreDecision(
                    shouldRestore: false,
                    reason: "sharedApplicationHidden",
                    diagnosticDetails: "previousWindowRestore=Skipped,reason=sharedApplicationHidden,whiteExcelBook1ExposureRisk=avoidSharedAppVisibleReexposure");
            }

            return new CaseWorkbookPreviousWindowRestoreDecision(
                shouldRestore: true,
                reason: "sharedApplicationVisible",
                diagnosticDetails: "previousWindowRestore=Required,reason=sharedApplicationVisible");
        }

        internal string BuildHiddenForDisplayStateCapturedMessage(
            CaseWorkbookPresentationHandoffPlan plan,
            long elapsedMilliseconds)
        {
            CaseWorkbookPresentationHandoffPlan safePlan = RequirePlan(plan);
            CaseWorkbookSharedDisplayStateFacts facts = safePlan.SharedStateFacts;
            return "Case workbook hidden-for-display Excel state captured. path="
                + safePlan.CaseWorkbookPath
                + ", route="
                + safePlan.RouteDecision.RouteName
                + ", "
                + safePlan.RouteDecision.ApplicationOwnerFacts
                + ", screenUpdating="
                + facts.PreviousScreenUpdating.ToString()
                + ", enableEvents="
                + facts.PreviousEnableEvents.ToString()
                + ", displayAlerts="
                + facts.PreviousDisplayAlerts.ToString()
                + ", elapsedMs="
                + elapsedMilliseconds.ToString();
        }

        internal string BuildHiddenForDisplayStateAppliedMessage(
            CaseWorkbookPresentationHandoffPlan plan,
            long elapsedMilliseconds)
        {
            CaseWorkbookPresentationHandoffPlan safePlan = RequirePlan(plan);
            return "Case workbook hidden-for-display Excel state applied. path="
                + safePlan.CaseWorkbookPath
                + ", route="
                + safePlan.RouteDecision.RouteName
                + ", "
                + safePlan.RouteDecision.ApplicationOwnerFacts
                + ", screenUpdating=false, enableEvents=false, displayAlerts=false, elapsedMs="
                + elapsedMilliseconds.ToString();
        }

        internal string BuildHiddenForDisplayOpenCompletedMessage(
            CaseWorkbookPresentationHandoffPlan plan,
            string appHwnd,
            long elapsedMilliseconds)
        {
            CaseWorkbookPresentationHandoffPlan safePlan = RequirePlan(plan);
            return "Case workbook hidden-for-display open completed. path="
                + safePlan.CaseWorkbookPath
                + ", route="
                + safePlan.RouteDecision.RouteName
                + ", appHwnd="
                + (appHwnd ?? string.Empty)
                + ", elapsedMs="
                + elapsedMilliseconds.ToString();
        }

        internal string BuildPreviousWindowRestoreSkippedMessage(
            CaseWorkbookPresentationHandoffPlan plan,
            long? elapsedMilliseconds)
        {
            CaseWorkbookPresentationHandoffPlan safePlan = RequirePlan(plan);
            return "Case workbook hidden-for-display previous window restore skipped because shared application was hidden. path="
                + safePlan.CaseWorkbookPath
                + ", route="
                + safePlan.RouteDecision.RouteName
                + ", elapsedMs="
                + (elapsedMilliseconds.HasValue ? elapsedMilliseconds.Value.ToString() : string.Empty);
        }

        internal string BuildSharedDisplayStateRestoredMessage(
            CaseWorkbookPresentationHandoffPlan plan,
            long? elapsedMilliseconds)
        {
            CaseWorkbookPresentationHandoffPlan safePlan = RequirePlan(plan);
            CaseWorkbookSharedDisplayStateFacts facts = safePlan.SharedStateFacts;
            return "Case workbook hidden Excel state restored. path="
                + safePlan.CaseWorkbookPath
                + ", route="
                + safePlan.RouteDecision.RouteName
                + ", "
                + safePlan.RouteDecision.ApplicationOwnerFacts
                + ", screenUpdating="
                + facts.PreviousScreenUpdating.ToString()
                + ", enableEvents="
                + facts.PreviousEnableEvents.ToString()
                + ", displayAlerts="
                + facts.PreviousDisplayAlerts.ToString()
                + ", elapsedMs="
                + (elapsedMilliseconds.HasValue ? elapsedMilliseconds.Value.ToString() : string.Empty);
        }

        internal string BuildSharedDisplayStateAppliedObservationDetails(CaseWorkbookPresentationHandoffPlan plan)
        {
            CaseWorkbookPresentationHandoffPlan safePlan = RequirePlan(plan);
            return "route=" + safePlan.RouteDecision.RouteName
                + "," + safePlan.RouteDecision.ApplicationOwnerFacts;
        }

        internal string BuildSharedDisplayStateRestoredObservationDetails(CaseWorkbookPresentationHandoffPlan plan)
        {
            return BuildSharedDisplayStateAppliedObservationDetails(plan);
        }

        internal string BuildVisibleOpenWindowFactsMessage(
            string stage,
            Excel.Application application,
            Excel.Workbook openedWorkbook)
        {
            CaseWorkbookPresentationVisibilityFacts facts = CaptureVisibilityFacts(application, openedWorkbook);
            return "Case workbook open visible state. stage="
                + (stage ?? string.Empty)
                + ", appHwnd="
                + facts.ApplicationHwnd
                + ", workbooksCount="
                + facts.WorkbooksCount
                + ", activeWorkbookName="
                + facts.ActiveWorkbookName
                + ", activeWindowCaption="
                + facts.ActiveWindowCaption
                + ", openedWorkbookWindows="
                + facts.OpenedWorkbookWindows;
        }

        internal CaseWorkbookPresentationVisibilityFacts CaptureVisibilityFacts(
            Excel.Application application,
            Excel.Workbook openedWorkbook)
        {
            return new CaseWorkbookPresentationVisibilityFacts(
                CaptureApplicationHwnd(application),
                SafeWorkbooksCount(application),
                SafeWorkbookName(application == null ? null : application.ActiveWorkbook),
                SafeWindowCaption(application == null ? null : application.ActiveWindow),
                DescribeWorkbookWindows(openedWorkbook));
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

        internal bool CaptureApplicationVisible(Excel.Application application)
        {
            try
            {
                return application != null && application.Visible;
            }
            catch
            {
                return false;
            }
        }

        private static CaseWorkbookPresentationHandoffPlan RequirePlan(CaseWorkbookPresentationHandoffPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return plan;
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
    }

    internal sealed class CaseWorkbookPresentationHandoffPlan
    {
        internal CaseWorkbookPresentationHandoffPlan(
            string caseWorkbookPath,
            CaseWorkbookOpenRouteDecision routeDecision,
            CaseWorkbookSharedDisplayStateFacts sharedStateFacts,
            CaseWorkbookPreviousWindowRestoreDecision previousWindowRestoreDecision,
            string diagnosticDetails)
        {
            CaseWorkbookPath = caseWorkbookPath ?? string.Empty;
            RouteDecision = routeDecision ?? throw new ArgumentNullException(nameof(routeDecision));
            SharedStateFacts = sharedStateFacts ?? CaseWorkbookSharedDisplayStateFacts.Empty;
            PreviousWindowRestoreDecision = previousWindowRestoreDecision
                ?? new CaseWorkbookPreviousWindowRestoreDecision(false, "unknown", string.Empty);
            DiagnosticDetails = diagnosticDetails ?? string.Empty;
        }

        internal string CaseWorkbookPath { get; }

        internal CaseWorkbookOpenRouteDecision RouteDecision { get; }

        internal CaseWorkbookSharedDisplayStateFacts SharedStateFacts { get; }

        internal CaseWorkbookPreviousWindowRestoreDecision PreviousWindowRestoreDecision { get; }

        internal string DiagnosticDetails { get; }
    }

    internal sealed class CaseWorkbookSharedDisplayStateFacts
    {
        internal static readonly CaseWorkbookSharedDisplayStateFacts Empty = new CaseWorkbookSharedDisplayStateFacts(
            previousActiveWindow: null,
            previousApplicationVisible: false,
            previousScreenUpdating: false,
            previousEnableEvents: false,
            previousDisplayAlerts: false);

        internal CaseWorkbookSharedDisplayStateFacts(
            Excel.Window previousActiveWindow,
            bool previousApplicationVisible,
            bool previousScreenUpdating,
            bool previousEnableEvents,
            bool previousDisplayAlerts)
        {
            PreviousActiveWindow = previousActiveWindow;
            PreviousApplicationVisible = previousApplicationVisible;
            PreviousScreenUpdating = previousScreenUpdating;
            PreviousEnableEvents = previousEnableEvents;
            PreviousDisplayAlerts = previousDisplayAlerts;
        }

        internal Excel.Window PreviousActiveWindow { get; }

        internal bool PreviousApplicationVisible { get; }

        internal bool PreviousScreenUpdating { get; }

        internal bool PreviousEnableEvents { get; }

        internal bool PreviousDisplayAlerts { get; }
    }

    internal sealed class CaseWorkbookPreviousWindowRestoreDecision
    {
        internal CaseWorkbookPreviousWindowRestoreDecision(
            bool shouldRestore,
            string reason,
            string diagnosticDetails)
        {
            ShouldRestore = shouldRestore;
            Reason = reason ?? string.Empty;
            DiagnosticDetails = diagnosticDetails ?? string.Empty;
        }

        internal bool ShouldRestore { get; }

        internal string Reason { get; }

        internal string DiagnosticDetails { get; }
    }

    internal sealed class CaseWorkbookPresentationVisibilityFacts
    {
        internal CaseWorkbookPresentationVisibilityFacts(
            string applicationHwnd,
            string workbooksCount,
            string activeWorkbookName,
            string activeWindowCaption,
            string openedWorkbookWindows)
        {
            ApplicationHwnd = applicationHwnd ?? string.Empty;
            WorkbooksCount = workbooksCount ?? string.Empty;
            ActiveWorkbookName = activeWorkbookName ?? string.Empty;
            ActiveWindowCaption = activeWindowCaption ?? string.Empty;
            OpenedWorkbookWindows = openedWorkbookWindows ?? string.Empty;
        }

        internal string ApplicationHwnd { get; }

        internal string WorkbooksCount { get; }

        internal string ActiveWorkbookName { get; }

        internal string ActiveWindowCaption { get; }

        internal string OpenedWorkbookWindows { get; }
    }
}
