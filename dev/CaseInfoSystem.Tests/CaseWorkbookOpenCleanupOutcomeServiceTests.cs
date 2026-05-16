using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class CaseWorkbookOpenCleanupOutcomeServiceTests
    {
        [Fact]
        public void CreateDedicatedHiddenSessionOutcome_WhenCleanupIsNotRequired_ReturnsNotRequiredOutcome()
        {
            var service = CreateService();

            CaseWorkbookOpenHiddenCleanupOutcome outcome = service.CreateDedicatedHiddenSessionOutcome(
                "owner",
                @"C:\Cases\not-required.xlsx",
                CaseWorkbookOpenRouteDecisionService.LegacyHiddenRouteName,
                CaseWorkbookOpenCleanupFacts.Empty,
                cleanupFailed: false);

            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.HiddenExcelCleanupNotRequired, outcome.HiddenCleanupOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.IsolatedAppReleaseNotRequired, outcome.IsolatedAppOutcome);
            Assert.Equal(string.Empty, outcome.RetainedInstanceOutcome);
            Assert.False(outcome.CacheReturnedToIdle);
            Assert.False(outcome.CachePoisoned);
            Assert.Contains("applicationKind=isolated", outcome.Details);
            Assert.Contains("isRetainedHiddenAppCache=False", outcome.Details);
            Assert.Contains("outcomeReason=dedicatedSessionFinalized", outcome.Details);
        }

        [Fact]
        public void CreateDedicatedHiddenSessionOutcome_WhenWorkbookAndAppCleanupSucceeded_ReturnsCompletedAndReleased()
        {
            var service = CreateService();

            CaseWorkbookOpenHiddenCleanupOutcome outcome = service.CreateDedicatedHiddenSessionOutcome(
                "owner",
                @"C:\Cases\completed.xlsx",
                CaseWorkbookOpenRouteDecisionService.ExperimentalIsolatedInnerSaveRouteName,
                new CaseWorkbookOpenCleanupFacts(
                    workbookPresent: true,
                    workbookCloseAttempted: true,
                    workbookCloseCompleted: true,
                    appPresent: true,
                    appQuitAttempted: true,
                    appQuitCompleted: true),
                cleanupFailed: false);

            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.HiddenExcelCleanupCompleted, outcome.HiddenCleanupOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.IsolatedAppReleased, outcome.IsolatedAppOutcome);
            Assert.Contains("route=experimental-isolated-inner-save", outcome.Details);
            Assert.Contains("workbookCloseCompleted=True", outcome.Details);
            Assert.Contains("appQuitCompleted=True", outcome.Details);
        }

        [Fact]
        public void CreateCachedHiddenSessionReturnedToIdleOutcome_WhenRetainedCacheIsKept_ReturnsReturnedToIdleTrace()
        {
            var service = CreateService();

            CaseWorkbookOpenHiddenCleanupOutcome outcome = service.CreateCachedHiddenSessionReturnedToIdleOutcome(
                "owner",
                @"C:\Cases\retained.xlsx",
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                new CaseWorkbookOpenCleanupFacts(
                    workbookPresent: true,
                    workbookCloseAttempted: true,
                    workbookCloseCompleted: true,
                    appPresent: true,
                    appQuitAttempted: false,
                    appQuitCompleted: false));

            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.HiddenExcelCleanupCompleted, outcome.HiddenCleanupOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.RetainedInstanceReturnedToIdle, outcome.RetainedInstanceOutcome);
            Assert.True(outcome.CacheReturnedToIdle);
            Assert.False(outcome.CachePoisoned);
            Assert.Contains("applicationKind=retained-hidden-app-cache", outcome.Details);
            Assert.Contains("cacheReturnedToIdle=True", outcome.Details);
            Assert.Contains("appQuitAttempted=False", outcome.Details);
            Assert.Contains("outcomeReason=returnedToIdle", outcome.Details);
        }

        [Fact]
        public void CreateCachedHiddenSessionPoisonedOutcome_WhenWorkbookCloseFailed_ReturnsDegradedAndPoisoned()
        {
            var service = CreateService();

            CaseWorkbookOpenHiddenCleanupOutcome outcome = service.CreateCachedHiddenSessionPoisonedOutcome(
                "owner",
                @"C:\Cases\poisoned.xlsx",
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                new CaseWorkbookOpenCleanupFacts(
                    workbookPresent: true,
                    workbookCloseAttempted: true,
                    workbookCloseCompleted: false,
                    appPresent: true,
                    appQuitAttempted: false,
                    appQuitCompleted: false),
                "workbookCloseFailed");

            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.HiddenExcelCleanupDegraded, outcome.HiddenCleanupOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.RetainedInstancePoisoned, outcome.RetainedInstanceOutcome);
            Assert.False(outcome.CacheReturnedToIdle);
            Assert.True(outcome.CachePoisoned);
            Assert.Contains("workbookCloseCompleted=False", outcome.Details);
            Assert.Contains("cachePoisoned=True", outcome.Details);
            Assert.Contains("outcomeReason=workbookCloseFailed", outcome.Details);
        }

        [Fact]
        public void CreateRetainedInstanceCleanupOutcome_ClassifiesNotRequiredSkippedCompletedAndDegraded()
        {
            var service = CreateService();

            CaseWorkbookOpenRetainedCleanupOutcome notRequired = service.CreateRetainedInstanceCleanupOutcome(
                "no-slot",
                appHwnd: string.Empty,
                retainedInstancePresent: false,
                isOwnedByCache: false,
                quitAttempted: false,
                quitCompleted: false);
            CaseWorkbookOpenRetainedCleanupOutcome skipped = service.CreateRetainedInstanceCleanupOutcome(
                "owner-mismatch",
                appHwnd: "42",
                retainedInstancePresent: true,
                isOwnedByCache: false,
                quitAttempted: false,
                quitCompleted: false);
            CaseWorkbookOpenRetainedCleanupOutcome completed = service.CreateRetainedInstanceCleanupOutcome(
                "shutdown-cleanup",
                appHwnd: "43",
                retainedInstancePresent: true,
                isOwnedByCache: true,
                quitAttempted: true,
                quitCompleted: true);
            CaseWorkbookOpenRetainedCleanupOutcome degraded = service.CreateRetainedInstanceCleanupOutcome(
                "idle-timeout",
                appHwnd: "44",
                retainedInstancePresent: true,
                isOwnedByCache: true,
                quitAttempted: true,
                quitCompleted: false);

            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.RetainedInstanceCleanupNotRequired, notRequired.RetainedInstanceOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.RetainedInstanceCleanupSkipped, skipped.RetainedInstanceOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.RetainedInstanceCleanupCompleted, completed.RetainedInstanceOutcome);
            Assert.Equal(CaseWorkbookOpenCleanupOutcomeService.RetainedInstanceCleanupDegraded, degraded.RetainedInstanceOutcome);
            Assert.Contains("cleanupReason=shutdown-cleanup", completed.KernelFlickerTraceMessage);
            Assert.Contains("appQuitCompleted=False", degraded.KernelFlickerTraceMessage);
        }

        [Fact]
        public void HiddenCleanupOutcome_BuildsKernelFlickerTraceAndObservationDetails()
        {
            var service = CreateService();

            CaseWorkbookOpenHiddenCleanupOutcome outcome = service.CreateCachedHiddenSessionReturnedToIdleOutcome(
                "CaseWorkbookOpenStrategy.CleanupCachedHiddenSession",
                @"C:\Cases\trace.xlsx",
                CaseWorkbookOpenRouteDecisionService.HiddenApplicationCacheRouteName,
                new CaseWorkbookOpenCleanupFacts(
                    workbookPresent: true,
                    workbookCloseAttempted: true,
                    workbookCloseCompleted: true,
                    appPresent: true,
                    appQuitAttempted: false,
                    appQuitCompleted: false));

            Assert.StartsWith(
                @"[KernelFlickerTrace] source=CaseWorkbookOpenStrategy.CleanupCachedHiddenSession action=hidden-excel-cleanup-outcome path=C:\Cases\trace.xlsx, scope=hidden-cleanup",
                outcome.KernelFlickerTraceMessage,
                StringComparison.Ordinal);
            Assert.Contains("hiddenCleanupOutcome=HiddenExcelCleanupCompleted", outcome.Details);
            Assert.Contains("retainedInstanceOutcome=RetainedInstanceReturnedToIdle", outcome.Details);
        }

        [Fact]
        public void Source_DoesNotOwnCleanupExecutionWorkbookCloseAppQuitOrComRelease()
        {
            string serviceSource = ReadInfrastructureSource("CaseWorkbookOpenCleanupOutcomeService.cs");
            string strategySource = ReadInfrastructureSource("CaseWorkbookOpenStrategy.cs");

            Assert.Contains("new CaseWorkbookOpenCleanupOutcomeService", strategySource);
            Assert.Contains("_cleanupOutcomeService.CreateDedicatedHiddenSessionOutcome", strategySource);
            Assert.Contains("_cleanupOutcomeService.CreateCachedHiddenSessionReturnedToIdleOutcome", strategySource);
            Assert.Contains("_cleanupOutcomeService.CreateCachedHiddenSessionPoisonedOutcome", strategySource);
            Assert.Contains("_cleanupOutcomeService.CreateRetainedInstanceCleanupOutcome", strategySource);
            Assert.Contains("CleanupDedicatedHiddenSession(", strategySource);
            Assert.Contains("CleanupCachedHiddenSession(", strategySource);
            Assert.Contains("DisposeCachedHiddenApplicationSlot(", strategySource);
            Assert.Contains("WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave", strategySource);
            Assert.Contains("TryQuitApplication(", strategySource);
            Assert.Contains("ReleaseComObject(", strategySource);

            Assert.DoesNotContain("Microsoft.Office.Interop.Excel", serviceSource);
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
            Assert.DoesNotContain("PrepareHiddenApplicationForUse", serviceSource);
            Assert.DoesNotContain("Visible =", serviceSource);
            Assert.DoesNotContain("WindowState", serviceSource);
            Assert.DoesNotContain("ScreenUpdating", serviceSource);
            Assert.DoesNotContain("Timer", serviceSource);
        }

        private static CaseWorkbookOpenCleanupOutcomeService CreateService()
        {
            return new CaseWorkbookOpenCleanupOutcomeService(new CaseWorkbookOpenRouteDecisionService());
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
