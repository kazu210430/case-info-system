using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public sealed class CaseWorkbookHiddenAppLifecycleSupportServiceTests
    {
        [Fact]
        public void CaptureLifecycleFacts_WhenApplicationIsReusable_ReturnsReusableFacts()
        {
            var service = new CaseWorkbookHiddenAppLifecycleSupportService();
            Excel.Application application = CreateReusableHiddenApplication();
            DateTime now = new DateTime(2026, 5, 16, 0, 0, 0, DateTimeKind.Utc);

            CaseWorkbookHiddenAppLifecycleFacts facts = service.CaptureLifecycleFacts(
                application,
                isInUse: false,
                isPoisoned: false,
                isOwnedByCache: true,
                idleSinceUtc: now.AddSeconds(-5),
                idleTimeoutSeconds: 60,
                utcNow: now);

            Assert.True(facts.IsReusable);
            Assert.Equal(string.Empty, facts.ReuseBlockReason);
            Assert.Equal("42", facts.AppHwnd);
            Assert.Equal(0, facts.WorkbooksCount);
            Assert.Contains("isReusable=True", facts.DiagnosticDetails);
            Assert.Contains("workbooksCount=0", facts.DiagnosticDetails);
            Assert.Contains("idleTimeoutExpired=False", facts.DiagnosticDetails);
        }

        [Fact]
        public void CaptureLifecycleFacts_WhenApplicationHasOpenWorkbookOrVisibleState_ReturnsNotReusableReason()
        {
            var service = new CaseWorkbookHiddenAppLifecycleSupportService();
            DateTime now = new DateTime(2026, 5, 16, 0, 0, 0, DateTimeKind.Utc);
            Excel.Application applicationWithWorkbook = CreateReusableHiddenApplication();
            applicationWithWorkbook.Workbooks.Add(new Excel.Workbook());
            Excel.Application visibleApplication = CreateReusableHiddenApplication();
            visibleApplication.Visible = true;

            CaseWorkbookHiddenAppLifecycleFacts workbookFacts = service.CaptureLifecycleFacts(
                applicationWithWorkbook,
                isInUse: false,
                isPoisoned: false,
                isOwnedByCache: true,
                idleSinceUtc: now,
                idleTimeoutSeconds: 60,
                utcNow: now);
            CaseWorkbookHiddenAppLifecycleFacts visibleFacts = service.CaptureLifecycleFacts(
                visibleApplication,
                isInUse: false,
                isPoisoned: false,
                isOwnedByCache: true,
                idleSinceUtc: now,
                idleTimeoutSeconds: 60,
                utcNow: now);
            CaseWorkbookHiddenAppLifecycleFacts poisonedFacts = service.CaptureLifecycleFacts(
                CreateReusableHiddenApplication(),
                isInUse: false,
                isPoisoned: true,
                isOwnedByCache: true,
                idleSinceUtc: now,
                idleTimeoutSeconds: 60,
                utcNow: now);

            Assert.False(workbookFacts.IsReusable);
            Assert.Equal(CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonWorkbooksOpen, workbookFacts.ReuseBlockReason);
            Assert.False(visibleFacts.IsReusable);
            Assert.Equal(CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonApplicationVisible, visibleFacts.ReuseBlockReason);
            Assert.False(poisonedFacts.IsReusable);
            Assert.Equal(CaseWorkbookHiddenAppLifecycleSupportService.ReuseBlockReasonPoisoned, poisonedFacts.ReuseBlockReason);
        }

        [Fact]
        public void DecideExpiration_ClassifiesIdlePoisonedAndExpiredStatesWithoutCleanupExecution()
        {
            var service = new CaseWorkbookHiddenAppLifecycleSupportService();
            DateTime now = new DateTime(2026, 5, 16, 0, 0, 0, DateTimeKind.Utc);

            CaseWorkbookHiddenAppExpirationDecision noSlot = service.DecideExpiration(
                service.CreateLifecycleStateFacts(
                    applicationPresent: false,
                    isInUse: false,
                    isPoisoned: false,
                    isOwnedByCache: false,
                    appHwnd: string.Empty,
                    idleSinceUtc: DateTime.MinValue,
                    idleTimeoutSeconds: 60,
                    utcNow: now),
                "idle-timeout");
            CaseWorkbookHiddenAppExpirationDecision inUse = service.DecideExpiration(
                service.CreateLifecycleStateFacts(
                    applicationPresent: true,
                    isInUse: true,
                    isPoisoned: false,
                    isOwnedByCache: true,
                    appHwnd: "42",
                    idleSinceUtc: DateTime.MinValue,
                    idleTimeoutSeconds: 60,
                    utcNow: now),
                "idle-timeout");
            CaseWorkbookHiddenAppExpirationDecision poisoned = service.DecideExpiration(
                service.CreateLifecycleStateFacts(
                    applicationPresent: true,
                    isInUse: false,
                    isPoisoned: true,
                    isOwnedByCache: true,
                    appHwnd: "42",
                    idleSinceUtc: now.AddSeconds(-5),
                    idleTimeoutSeconds: 60,
                    utcNow: now),
                "idle-timeout");
            CaseWorkbookHiddenAppExpirationDecision initializeIdle = service.DecideExpiration(
                service.CreateLifecycleStateFacts(
                    applicationPresent: true,
                    isInUse: false,
                    isPoisoned: false,
                    isOwnedByCache: true,
                    appHwnd: "42",
                    idleSinceUtc: DateTime.MinValue,
                    idleTimeoutSeconds: 60,
                    utcNow: now),
                "idle-timeout");
            CaseWorkbookHiddenAppExpirationDecision notExpired = service.DecideExpiration(
                service.CreateLifecycleStateFacts(
                    applicationPresent: true,
                    isInUse: false,
                    isPoisoned: false,
                    isOwnedByCache: true,
                    appHwnd: "42",
                    idleSinceUtc: now.AddSeconds(-30),
                    idleTimeoutSeconds: 60,
                    utcNow: now),
                "idle-timeout");
            CaseWorkbookHiddenAppExpirationDecision expired = service.DecideExpiration(
                service.CreateLifecycleStateFacts(
                    applicationPresent: true,
                    isInUse: false,
                    isPoisoned: false,
                    isOwnedByCache: true,
                    appHwnd: "42",
                    idleSinceUtc: now.AddSeconds(-61),
                    idleTimeoutSeconds: 60,
                    utcNow: now),
                "idle-timeout");

            Assert.False(noSlot.DisposeSlot);
            Assert.True(noSlot.StopIdleTimer);
            Assert.False(inUse.DisposeSlot);
            Assert.True(inUse.StopIdleTimer);
            Assert.True(poisoned.DisposeSlot);
            Assert.Equal("poisoned", poisoned.DecisionReason);
            Assert.False(initializeIdle.DisposeSlot);
            Assert.True(initializeIdle.InitializeIdleSinceUtc);
            Assert.Equal(now, initializeIdle.InitializedIdleSinceUtc);
            Assert.False(notExpired.DisposeSlot);
            Assert.Equal("idle-timeout-not-reached", notExpired.DecisionReason);
            Assert.True(expired.DisposeSlot);
            Assert.Equal("idle-timeout-reached", expired.DecisionReason);
        }

        [Fact]
        public void BuildTraceMessages_PreservesRetainedCacheReasonAndDiagnosticTerms()
        {
            var service = new CaseWorkbookHiddenAppLifecycleSupportService();
            var routeDecisionService = new CaseWorkbookOpenRouteDecisionService();
            CaseWorkbookOpenRouteDecision fallbackDecision =
                routeDecisionService.DecideHiddenApplicationCacheAcquisition(cachedApplicationInUse: true);

            string bypassMessage = service.BuildCacheBypassInUseMessage(@"C:\Cases\bypass.xlsx", fallbackDecision, 12);
            string returnedMessage = service.BuildReturnedToIdleMessage(@"C:\Cases\idle.xlsx", "app-cache", "42", 60, 13);
            string skippedMessage = service.BuildCleanupSkippedNotOwnedMessage("owner-mismatch", "43");
            string discardedMessage = service.BuildDiscardedMessage("shutdown-cleanup", "44");
            string details = service.BuildAcquiredObservationDetails(
                "app-cache",
                reusedApplication: true,
                "applicationKind=retained-hidden-app-cache");
            string acquireEvent = service.BuildDiagnosticEventMessage(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionAcquire,
                @"C:\Cases\idle.xlsx",
                "app-cache",
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                "cache-reusable")
            {
                EventOutcome = "acquired",
                CacheEvent = "acquire",
                AcquisitionKind = "reused",
                ReusedApplication = true,
                AppHwnd = "42",
                ApplicationOwnerFacts = "applicationKind=retained-hidden-app-cache"
            });
            string fallbackEvent = service.BuildDiagnosticEventMessage(new CaseWorkbookHiddenAppLifecycleDiagnosticEvent(
                CaseWorkbookHiddenAppLifecycleSupportService.LifecycleActionFallback,
                @"C:\Cases\bypass.xlsx",
                "app-cache-bypass-inuse",
                "CaseWorkbookOpenStrategy.OpenHiddenWorkbookWithApplicationCache",
                "hiddenApplicationCacheInUse")
            {
                EventOutcome = "fallback-to-dedicated-hidden-session",
                CacheEvent = "acquire-fallback",
                FallbackRoute = "app-cache-bypass-inuse",
                AbandonedOperation = "retained-cache-acquire",
                SafetyAction = "open-dedicated-hidden-session"
            });

            Assert.Contains("bypassed because in-use", bypassMessage);
            Assert.Contains("route=app-cache-bypass-inuse", bypassMessage);
            Assert.Contains("routeReason=hiddenApplicationCacheInUse", bypassMessage);
            Assert.Contains("returned-to-idle", returnedMessage);
            Assert.Contains("idleTimeoutSeconds=60", returnedMessage);
            Assert.Contains("cleanup skipped because slot is not cache-owned", skippedMessage);
            Assert.Contains("reason=owner-mismatch", skippedMessage);
            Assert.Contains("discarded", discardedMessage);
            Assert.Contains("reason=shutdown-cleanup", discardedMessage);
            Assert.Equal("route=app-cache,reused=True,applicationKind=retained-hidden-app-cache", details);
            Assert.Contains("action=retained-hidden-app-cache-acquire", acquireEvent);
            Assert.Contains("acquisitionKind=reused", acquireEvent);
            Assert.Contains("reusedApplication=True", acquireEvent);
            Assert.Contains("action=retained-hidden-app-cache-fallback", fallbackEvent);
            Assert.Contains("abandonedOperation=retained-cache-acquire", fallbackEvent);
            Assert.Contains("safetyAction=open-dedicated-hidden-session", fallbackEvent);
        }

        [Fact]
        public void SourceBoundary_HelperDoesNotOwnTimerPoisonShutdownCloseQuitOrComRelease()
        {
            string serviceSource = ReadInfrastructureSource("CaseWorkbookHiddenAppLifecycleSupportService.cs");
            string strategySource = ReadInfrastructureSource("CaseWorkbookOpenStrategy.cs");

            Assert.Contains("new CaseWorkbookHiddenAppLifecycleSupportService", strategySource);
            Assert.Contains("_hiddenAppLifecycleSupportService.CaptureLifecycleFacts", strategySource);
            Assert.Contains("HiddenApplicationIdleTimer_Tick", strategySource);
            Assert.Contains("DisposeHiddenApplicationIdleTimerUnlocked", strategySource);
            Assert.Contains("MarkCachedHiddenApplicationPoisoned", strategySource);
            Assert.Contains("DisposeCachedHiddenApplicationSlot", strategySource);
            Assert.Contains("WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave", strategySource);
            Assert.Contains("TryQuitApplication(", strategySource);
            Assert.Contains("ReleaseComObject(", strategySource);

            Assert.DoesNotContain("WorkbookCloseInteropHelper", serviceSource);
            Assert.DoesNotContain("CloseOwnedWorkbook", serviceSource);
            Assert.DoesNotContain(".Close(", serviceSource);
            Assert.DoesNotContain(".Quit(", serviceSource);
            Assert.DoesNotContain("TryQuitApplication", serviceSource);
            Assert.DoesNotContain("ReleaseComObject", serviceSource);
            Assert.DoesNotContain("FinalReleaseComObject", serviceSource);
            Assert.DoesNotContain("Marshal.ReleaseComObject", serviceSource);
            Assert.DoesNotContain("HiddenApplicationIdleTimer_Tick", serviceSource);
            Assert.DoesNotContain("DisposeHiddenApplicationIdleTimer", serviceSource);
            Assert.DoesNotContain("StopHiddenApplicationIdleTimer", serviceSource);
            Assert.DoesNotContain("MarkCachedHiddenApplicationPoisoned", serviceSource);
            Assert.DoesNotContain("DisposeCachedHiddenApplicationSlot", serviceSource);
            Assert.DoesNotContain("ShutdownHiddenApplicationCache", serviceSource);
            Assert.DoesNotContain("PrepareHiddenApplicationForUse", serviceSource);
            Assert.DoesNotContain("Workbooks.Open", serviceSource);
            Assert.DoesNotContain(".Visible =", serviceSource);
            Assert.DoesNotContain(".ScreenUpdating =", serviceSource);
            Assert.DoesNotContain(".EnableEvents =", serviceSource);
            Assert.DoesNotContain(".DisplayAlerts =", serviceSource);
            Assert.DoesNotContain(".UserControl =", serviceSource);
        }

        private static Excel.Application CreateReusableHiddenApplication()
        {
            return new Excel.Application
            {
                Hwnd = 42,
                Visible = false,
                Ready = true,
                DisplayAlerts = false,
                ScreenUpdating = false,
                EnableEvents = false,
                UserControl = false
            };
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
                if (File.Exists(Path.Combine(current.FullName, "build.ps1")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new InvalidOperationException("Repository root was not found.");
        }
    }
}
