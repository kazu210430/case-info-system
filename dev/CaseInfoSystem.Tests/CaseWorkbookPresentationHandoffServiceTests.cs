using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    [Collection("CaseWorkbookOpenStrategy")]
    public sealed class CaseWorkbookPresentationHandoffServiceTests
    {
        [Fact]
        public void CreateHiddenForDisplayPlan_WhenSharedApplicationIsVisible_BuildsRestoreDecisionAndDiagnostics()
        {
            var service = new CaseWorkbookPresentationHandoffService();
            CaseWorkbookOpenRouteDecision routeDecision =
                new CaseWorkbookOpenRouteDecisionService().DecideCreatedCaseDisplayRoute();
            var previousWindow = new Excel.Window { Visible = true, Hwnd = 10 };

            CaseWorkbookPresentationHandoffPlan plan = service.CreateHiddenForDisplayPlan(
                @"C:\Cases\display.xlsx",
                routeDecision,
                previousWindow,
                previousApplicationVisible: true,
                previousScreenUpdating: true,
                previousEnableEvents: false,
                previousDisplayAlerts: true);

            Assert.Equal(CaseWorkbookOpenRouteDecisionService.CreatedCaseDisplayHiddenRouteName, plan.RouteDecision.RouteName);
            Assert.Same(previousWindow, plan.SharedStateFacts.PreviousActiveWindow);
            Assert.True(plan.SharedStateFacts.PreviousApplicationVisible);
            Assert.True(plan.SharedStateFacts.PreviousScreenUpdating);
            Assert.False(plan.SharedStateFacts.PreviousEnableEvents);
            Assert.True(plan.SharedStateFacts.PreviousDisplayAlerts);
            Assert.True(plan.PreviousWindowRestoreDecision.ShouldRestore);
            Assert.Equal("sharedApplicationVisible", plan.PreviousWindowRestoreDecision.Reason);
            Assert.Contains("scope=presentation-handoff", plan.DiagnosticDetails);
            Assert.Contains("applicationKind=shared-current", plan.DiagnosticDetails);
            Assert.Contains("previousWindowRestoreRequired=True", plan.DiagnosticDetails);
        }

        [Fact]
        public void CreateHiddenForDisplayPlan_WhenSharedApplicationWasHidden_SkipsPreviousWindowRestoreForWhiteExcelExposureBoundary()
        {
            var service = new CaseWorkbookPresentationHandoffService();
            CaseWorkbookOpenRouteDecision routeDecision =
                new CaseWorkbookOpenRouteDecisionService().DecideCreatedCaseDisplayRoute();
            var previousWindow = new Excel.Window { Visible = false, Hwnd = 20 };

            CaseWorkbookPresentationHandoffPlan plan = service.CreateHiddenForDisplayPlan(
                @"C:\Cases\hidden-app.xlsx",
                routeDecision,
                previousWindow,
                previousApplicationVisible: false,
                previousScreenUpdating: true,
                previousEnableEvents: true,
                previousDisplayAlerts: true);

            Assert.False(plan.PreviousWindowRestoreDecision.ShouldRestore);
            Assert.Equal("sharedApplicationHidden", plan.PreviousWindowRestoreDecision.Reason);
            Assert.Contains(
                "whiteExcelBook1ExposureRisk=avoidSharedAppVisibleReexposure",
                plan.PreviousWindowRestoreDecision.DiagnosticDetails);
            Assert.Contains("previousApplicationVisible=False", plan.DiagnosticDetails);
            Assert.Contains("previousWindowRestoreRequired=False", plan.DiagnosticDetails);
            Assert.Equal(
                @"Case workbook hidden-for-display previous window restore skipped because shared application was hidden. path=C:\Cases\hidden-app.xlsx, route=created-case-display, elapsedMs=42",
                service.BuildPreviousWindowRestoreSkippedMessage(plan, 42));
        }

        [Fact]
        public void BuildHiddenForDisplayMessages_PreservesExistingRouteOwnerAndStateFacts()
        {
            var service = new CaseWorkbookPresentationHandoffService();
            CaseWorkbookOpenRouteDecision routeDecision =
                new CaseWorkbookOpenRouteDecisionService().DecideCreatedCaseDisplayRoute();
            CaseWorkbookPresentationHandoffPlan plan = service.CreateHiddenForDisplayPlan(
                @"C:\Cases\display.xlsx",
                routeDecision,
                new Excel.Window { Visible = true },
                previousApplicationVisible: true,
                previousScreenUpdating: true,
                previousEnableEvents: false,
                previousDisplayAlerts: true);

            string captured = service.BuildHiddenForDisplayStateCapturedMessage(plan, 12);
            string applied = service.BuildHiddenForDisplayStateAppliedMessage(plan, 13);
            string restored = service.BuildSharedDisplayStateRestoredMessage(plan, 14);

            Assert.Contains("route=created-case-display", captured);
            Assert.Contains("applicationLifetimeOwner=user-or-excel-host", captured);
            Assert.Contains("screenUpdating=True", captured);
            Assert.Contains("enableEvents=False", captured);
            Assert.Contains("displayAlerts=True", captured);
            Assert.Contains("screenUpdating=false, enableEvents=false, displayAlerts=false", applied);
            Assert.Contains("screenUpdating=True", restored);
            Assert.Contains("enableEvents=False", restored);
            Assert.Contains("displayAlerts=True", restored);
            Assert.Equal(
                "route=created-case-display,applicationKind=shared-current,applicationLifetimeOwner=user-or-excel-host,isSharedCurrentApp=True,isIsolatedApp=False,isRetainedHiddenAppCache=False",
                service.BuildSharedDisplayStateAppliedObservationDetails(plan));
        }

        [Fact]
        public void BuildVisibleOpenWindowFactsMessage_CapturesApplicationWorkbookAndWindowFacts()
        {
            var service = new CaseWorkbookPresentationHandoffService();
            var application = new Excel.Application { Hwnd = 100 };
            var activeWorkbook = new Excel.Workbook { Name = "ActiveBook.xlsx" };
            application.Workbooks.Add(activeWorkbook);
            application.ActiveWorkbook = activeWorkbook;
            application.ActiveWindow = new Excel.Window { Caption = "Active Window", Hwnd = 200, Visible = true };

            var openedWorkbook = new Excel.Workbook { Name = "Display.xlsx" };
            openedWorkbook.Windows.Add(new Excel.Window { Visible = false, Caption = "Hidden Case", Hwnd = 301 });
            openedWorkbook.Windows.Add(new Excel.Window { Visible = true, Caption = "Visible Case", Hwnd = 302 });
            application.Workbooks.Add(openedWorkbook);

            string message = service.BuildVisibleOpenWindowFactsMessage("after-open", application, openedWorkbook);

            Assert.Contains("stage=after-open", message);
            Assert.Contains("appHwnd=100", message);
            Assert.Contains("workbooksCount=2", message);
            Assert.Contains("activeWorkbookName=ActiveBook.xlsx", message);
            Assert.Contains("activeWindowCaption=Active Window", message);
            Assert.Contains("openedWorkbookWindows=count=2", message);
            Assert.Contains("index=1,visible=False,caption=Hidden Case,hwnd=301", message);
            Assert.Contains("index=2,visible=True,caption=Visible Case,hwnd=302", message);
        }

        [Fact]
        public void Source_DoesNotOwnPresentationExecutionWindowMutationCleanupOrLifecycle()
        {
            string serviceSource = ReadInfrastructureSource("CaseWorkbookPresentationHandoffService.cs");
            string strategySource = ReadInfrastructureSource("CaseWorkbookOpenStrategy.cs");

            Assert.Contains("new CaseWorkbookPresentationHandoffService()", strategySource);
            Assert.Contains("_presentationHandoffService.CreateHiddenForDisplayPlan", strategySource);
            Assert.Contains("RestorePreviousWindow(", strategySource);
            Assert.Contains("HideOpenedWorkbookWindow(", strategySource);
            Assert.Contains("_application.Workbooks.Open", strategySource);
            Assert.Contains("WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave", strategySource);
            Assert.Contains("TryQuitApplication(", strategySource);
            Assert.Contains("ReleaseComObject(", strategySource);

            Assert.Contains("Microsoft.Office.Interop.Excel", serviceSource);
            Assert.DoesNotContain("Workbooks.Open", serviceSource);
            Assert.DoesNotContain("WorkbookCloseInteropHelper", serviceSource);
            Assert.DoesNotContain("CloseOwnedWorkbook", serviceSource);
            Assert.DoesNotContain(".Close(", serviceSource);
            Assert.DoesNotContain(".Quit(", serviceSource);
            Assert.DoesNotContain("TryQuitApplication", serviceSource);
            Assert.DoesNotContain("ReleaseComObject", serviceSource);
            Assert.DoesNotContain("FinalReleaseComObject", serviceSource);
            Assert.DoesNotContain("Marshal.ReleaseComObject", serviceSource);
            Assert.DoesNotContain("HideOpenedWorkbookWindow", serviceSource);
            Assert.DoesNotContain("RestorePreviousWindow(", serviceSource);
            Assert.DoesNotContain("RestoreSharedApplicationState", serviceSource);
            Assert.DoesNotContain(".Visible =", serviceSource);
            Assert.DoesNotContain("application.Visible =", serviceSource);
            Assert.DoesNotContain("WindowState", serviceSource);
            Assert.DoesNotContain(".Activate(", serviceSource);
            Assert.DoesNotContain("Application.Visible", serviceSource);
        }

        private static string ReadInfrastructureSource(string infrastructureFileName)
        {
            string repoRoot = FindRepositoryRoot();
            return File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "Infrastructure", infrastructureFileName));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
