using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    [Collection("CaseWorkbookOpenStrategy")]
    public sealed class CaseWorkbookOpenRouteDecisionServiceTests
    {
        [Fact]
        public void DecideHiddenCreateRoute_WhenNoSwitchesAreEnabled_ReturnsLegacyIsolatedRoute()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                var service = new CaseWorkbookOpenRouteDecisionService();

                CaseWorkbookOpenRouteDecision decision = service.DecideHiddenCreateRoute();

                Assert.Equal(CaseWorkbookOpenRouteDecisionService.LegacyHiddenRouteName, decision.RouteName);
                Assert.Equal("defaultLegacyHiddenRoute", decision.Reason);
                Assert.False(decision.SaveBeforeClose);
                Assert.False(decision.UseHiddenApplicationCache);
                Assert.False(decision.IsFallbackRoute);
                Assert.True(decision.IsIsolatedApplication);
                Assert.False(decision.IsSharedCurrentApplication);
                Assert.False(decision.IsRetainedHiddenApplicationCache);
                Assert.Equal(CaseWorkbookOpenRouteDecisionService.ApplicationKindIsolated, decision.ApplicationKind);
                Assert.Equal(
                    CaseWorkbookOpenRouteDecisionService.ApplicationLifetimeOwnerCaseWorkbookOpenStrategy,
                    decision.ApplicationLifetimeOwner);
                Assert.Equal(
                    "route=legacy-isolated,routeReason=defaultLegacyHiddenRoute,saveBeforeClose=False,useHiddenApplicationCache=False,isFallbackRoute=False,applicationKind=isolated,applicationLifetimeOwner=CaseWorkbookOpenStrategy,isSharedCurrentApp=False,isIsolatedApp=True,isRetainedHiddenAppCache=False",
                    decision.RouteTraceDetails);
            }
        }

        [Fact]
        public void DecideHiddenCreateRoute_WhenHiddenApplicationCacheIsEnabled_ReturnsRetainedCacheRoute()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheEnvironmentVariableName,
                    "1");
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.DedicatedHiddenInnerSaveEnvironmentVariableName,
                    "1");
                var service = new CaseWorkbookOpenRouteDecisionService();

                CaseWorkbookOpenRouteDecision decision = service.DecideHiddenCreateRoute();

                Assert.Equal(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName, decision.RouteName);
                Assert.Equal("hiddenApplicationCacheEnabled", decision.Reason);
                Assert.True(decision.UseHiddenApplicationCache);
                Assert.False(decision.SaveBeforeClose);
                Assert.True(decision.IsRetainedHiddenApplicationCache);
                Assert.False(decision.IsIsolatedApplication);
                Assert.Equal(
                    "applicationKind=retained-hidden-app-cache,applicationLifetimeOwner=CaseWorkbookOpenStrategy,isSharedCurrentApp=False,isIsolatedApp=False,isRetainedHiddenAppCache=True",
                    decision.ApplicationOwnerFacts);
            }
        }

        [Fact]
        public void DecideHiddenCreateRoute_WhenDedicatedInnerSaveSwitchIsEnabled_ReturnsInnerSaveRoute()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.DedicatedHiddenInnerSaveEnvironmentVariableName,
                    "true");
                var service = new CaseWorkbookOpenRouteDecisionService();

                CaseWorkbookOpenRouteDecision decision = service.DecideHiddenCreateRoute();

                Assert.Equal(CaseWorkbookOpenRouteDecisionService.ExperimentalIsolatedInnerSaveRouteName, decision.RouteName);
                Assert.Equal("dedicatedHiddenInnerSaveEnabled", decision.Reason);
                Assert.True(decision.SaveBeforeClose);
                Assert.False(decision.UseHiddenApplicationCache);
                Assert.True(decision.IsIsolatedApplication);
            }
        }

        [Fact]
        public void DecideHiddenCreateRoute_WhenLegacyAliasSwitchIsEnabled_ReturnsInnerSaveRouteWithAliasReason()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName,
                    "1");
                var service = new CaseWorkbookOpenRouteDecisionService();

                CaseWorkbookOpenRouteDecision decision = service.DecideHiddenCreateRoute();

                Assert.Equal(CaseWorkbookOpenRouteDecisionService.ExperimentalIsolatedInnerSaveRouteName, decision.RouteName);
                Assert.Equal("legacyDedicatedHiddenInnerSaveAliasEnabled", decision.Reason);
                Assert.True(decision.SaveBeforeClose);
            }
        }

        [Fact]
        public void DecideHiddenApplicationCacheAcquisition_WhenCacheIsInUse_ReturnsBypassFallbackRoute()
        {
            var service = new CaseWorkbookOpenRouteDecisionService();

            CaseWorkbookOpenRouteDecision decision = service.DecideHiddenApplicationCacheAcquisition(cachedApplicationInUse: true);

            Assert.Equal(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheBypassInUseRouteName, decision.RouteName);
            Assert.Equal("hiddenApplicationCacheInUse", decision.Reason);
            Assert.True(decision.IsFallbackRoute);
            Assert.False(decision.UseHiddenApplicationCache);
            Assert.False(decision.SaveBeforeClose);
            Assert.True(decision.IsIsolatedApplication);
            Assert.Contains("route=app-cache-bypass-inuse", decision.RouteTraceDetails);
            Assert.Contains("isFallbackRoute=True", decision.RouteTraceDetails);
        }

        [Fact]
        public void DecideCreatedCaseDisplayRoute_ReturnsSharedCurrentApplicationFacts()
        {
            var service = new CaseWorkbookOpenRouteDecisionService();

            CaseWorkbookOpenRouteDecision decision = service.DecideCreatedCaseDisplayRoute();

            Assert.Equal(CaseWorkbookOpenRouteDecisionService.CreatedCaseDisplayHiddenRouteName, decision.RouteName);
            Assert.Equal("createdCaseDisplayHandoff", decision.Reason);
            Assert.True(decision.IsSharedCurrentApplication);
            Assert.False(decision.IsIsolatedApplication);
            Assert.False(decision.IsRetainedHiddenApplicationCache);
            Assert.Equal(CaseWorkbookOpenRouteDecisionService.ApplicationKindSharedCurrent, decision.ApplicationKind);
            Assert.Equal(
                CaseWorkbookOpenRouteDecisionService.ApplicationLifetimeOwnerUserOrExcelHost,
                decision.ApplicationLifetimeOwner);
            Assert.Equal(
                "applicationKind=shared-current,applicationLifetimeOwner=user-or-excel-host,isSharedCurrentApp=True,isIsolatedApp=False,isRetainedHiddenAppCache=False",
                decision.ApplicationOwnerFacts);
        }

        [Fact]
        public void ReadHiddenRouteSwitches_InterpretsTruthyValuesAndIdleSeconds()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.DedicatedHiddenInnerSaveEnvironmentVariableName,
                    "true");
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName,
                    "1");
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheEnvironmentVariableName,
                    "TRUE");
                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName,
                    "42");
                var service = new CaseWorkbookOpenRouteDecisionService();

                CaseWorkbookOpenRouteSwitches switches = service.ReadHiddenRouteSwitches();

                Assert.True(switches.DedicatedHiddenInnerSaveEnabled);
                Assert.True(switches.LegacyDedicatedHiddenInnerSaveAliasEnabled);
                Assert.True(switches.HiddenApplicationCacheEnabled);
                Assert.Equal(42, switches.HiddenApplicationCacheIdleSeconds);
                Assert.Equal("true", switches.DedicatedHiddenInnerSaveRawValue);
                Assert.Equal("1", switches.LegacyDedicatedHiddenInnerSaveAliasRawValue);
                Assert.Equal("TRUE", switches.HiddenApplicationCacheRawValue);
                Assert.Equal("42", switches.HiddenApplicationCacheIdleSecondsRawValue);
            }
        }

        [Fact]
        public void ReadHiddenRouteSwitches_UsesDefaultIdleSecondsForInvalidOrNonPositiveValues()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                var service = new CaseWorkbookOpenRouteDecisionService();

                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName,
                    "0");
                Assert.Equal(
                    CaseWorkbookOpenRouteDecisionService.DefaultHiddenApplicationCacheIdleSeconds,
                    service.ResolveHiddenApplicationCacheIdleSeconds());

                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName,
                    "-1");
                Assert.Equal(
                    CaseWorkbookOpenRouteDecisionService.DefaultHiddenApplicationCacheIdleSeconds,
                    service.ResolveHiddenApplicationCacheIdleSeconds());

                Environment.SetEnvironmentVariable(
                    CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName,
                    "invalid");
                Assert.Equal(
                    CaseWorkbookOpenRouteDecisionService.DefaultHiddenApplicationCacheIdleSeconds,
                    service.ResolveHiddenApplicationCacheIdleSeconds());
            }
        }

        [Fact]
        public void Source_DoesNotOwnExcelApplicationWorkbookOrLifecycleExecution()
        {
            string source = ReadInfrastructureSource("CaseWorkbookOpenRouteDecisionService.cs");

            Assert.DoesNotContain("Microsoft.Office.Interop.Excel", source);
            Assert.DoesNotContain("new Excel.Application", source);
            Assert.DoesNotContain("Workbooks.Open", source);
            Assert.DoesNotContain("WorkbookCloseInteropHelper", source);
            Assert.DoesNotContain("CloseOwnedWorkbook", source);
            Assert.DoesNotContain(".Close(", source);
            Assert.DoesNotContain(".Quit(", source);
            Assert.DoesNotContain("TryQuitApplication", source);
            Assert.DoesNotContain("ReleaseComObject", source);
            Assert.DoesNotContain("FinalReleaseComObject", source);
            Assert.DoesNotContain("Marshal.ReleaseComObject", source);
            Assert.DoesNotContain("HideOpenedWorkbookWindow", source);
            Assert.DoesNotContain("PrepareHiddenApplicationForUse", source);
            Assert.DoesNotContain("RestoreSharedApplicationState", source);
            Assert.DoesNotContain("Cleanup", source);
            Assert.DoesNotContain("Visible =", source);
            Assert.DoesNotContain("WindowState", source);
            Assert.DoesNotContain("ScreenUpdating", source);
            Assert.DoesNotContain("Timer", source);
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

        private sealed class HiddenRouteEnvironmentScope : IDisposable
        {
            private readonly string _dedicatedHiddenInnerSave;
            private readonly string _legacyDedicatedHiddenInnerSaveAlias;
            private readonly string _hiddenApplicationCache;
            private readonly string _hiddenApplicationCacheIdleSeconds;

            internal HiddenRouteEnvironmentScope()
            {
                _dedicatedHiddenInnerSave = Environment.GetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.DedicatedHiddenInnerSaveEnvironmentVariableName);
                _legacyDedicatedHiddenInnerSaveAlias = Environment.GetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName);
                _hiddenApplicationCache = Environment.GetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheEnvironmentVariableName);
                _hiddenApplicationCacheIdleSeconds = Environment.GetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.DedicatedHiddenInnerSaveEnvironmentVariableName, null);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName, null);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheEnvironmentVariableName, null);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName, null);
            }

            public void Dispose()
            {
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.DedicatedHiddenInnerSaveEnvironmentVariableName, _dedicatedHiddenInnerSave);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName, _legacyDedicatedHiddenInnerSaveAlias);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheEnvironmentVariableName, _hiddenApplicationCache);
                Environment.SetEnvironmentVariable(CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheIdleSecondsEnvironmentVariableName, _hiddenApplicationCacheIdleSeconds);
            }
        }
    }
}
