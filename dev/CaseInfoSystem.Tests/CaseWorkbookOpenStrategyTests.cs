using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    [CollectionDefinition("CaseWorkbookOpenStrategy", DisableParallelization = true)]
    public sealed class CaseWorkbookOpenStrategyCollection
    {
    }

    [Collection("CaseWorkbookOpenStrategy")]
    public class CaseWorkbookOpenStrategyTests
    {
        private const string DedicatedHiddenInnerSaveEnvironmentVariableName = "CASEINFO_EXPERIMENT_DEDICATED_HIDDEN_INNER_SAVE";
        private const string LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName = "CASEINFO_EXPERIMENT_SHARED_HIDDEN_EXCEL";
        private const string HiddenApplicationCacheEnvironmentVariableName = "CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE";
        private const string HiddenApplicationCacheIdleSecondsEnvironmentVariableName = "CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE_IDLE_SECONDS";

        [Fact]
        public void OpenHiddenWorkbook_UsesLegacyRoute_WhenNoExperimentalFlagsAreEnabled()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application hiddenApplication = CreateHiddenApplication();
                var strategy = CreateStrategy(logs, releasedObjects, hiddenApplication);

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession session = strategy.OpenHiddenWorkbook(@"C:\Cases\legacy.xlsx");
                Excel.Workbook workbook = session.Workbook;

                Assert.Equal("legacy-isolated", session.RouteName);
                Assert.Same(hiddenApplication, session.Application);

                session.Close();

                Assert.Equal(0, workbook.SaveCallCount);
                Assert.Equal(1, workbook.CloseCallCount);
                Assert.Equal(1, hiddenApplication.QuitCallCount);
                Assert.Contains(workbook, releasedObjects);
                Assert.Contains(hiddenApplication, releasedObjects);
            }
        }

        [Fact]
        public void OpenHiddenWorkbook_UsesIsolatedInnerSaveRoute_WhenDedicatedHiddenInnerSaveFlagIsEnabled()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(DedicatedHiddenInnerSaveEnvironmentVariableName, "1");
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application hiddenApplication = CreateHiddenApplication();
                var strategy = CreateStrategy(logs, releasedObjects, hiddenApplication);

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession session = strategy.OpenHiddenWorkbook(@"C:\Cases\shared.xlsx");
                Excel.Workbook workbook = session.Workbook;

                Assert.Equal("experimental-isolated-inner-save", session.RouteName);

                session.Close();

                Assert.Equal(1, workbook.SaveCallCount);
                Assert.Equal(1, workbook.CloseCallCount);
                Assert.Equal(1, hiddenApplication.QuitCallCount);
                Assert.Contains(workbook, releasedObjects);
                Assert.Contains(hiddenApplication, releasedObjects);
            }
        }

        [Fact]
        public void OpenHiddenWorkbook_UsesIsolatedInnerSaveRoute_WhenLegacyAliasFlagIsEnabled()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName, "1");
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application hiddenApplication = CreateHiddenApplication();
                var strategy = CreateStrategy(logs, releasedObjects, hiddenApplication);

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession session = strategy.OpenHiddenWorkbook(@"C:\Cases\legacy-alias.xlsx");

                Assert.Equal("experimental-isolated-inner-save", session.RouteName);

                session.Close();
            }
        }

        [Fact]
        public void OpenHiddenWorkbook_WithCacheEnabled_ReusesApplicationAcrossSessionsUntilHiddenApplicationCacheShutdown()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName, "1");
                Environment.SetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName, "60");
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application cachedApplication = CreateHiddenApplication();
                var strategy = CreateStrategy(logs, releasedObjects, cachedApplication);

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession firstSession = strategy.OpenHiddenWorkbook(@"C:\Cases\cached-1.xlsx");
                Excel.Workbook firstWorkbook = firstSession.Workbook;
                firstSession.Close();

                Assert.Equal("app-cache", firstSession.RouteName);
                Assert.Equal(1, firstWorkbook.CloseCallCount);
                Assert.Equal(0, cachedApplication.QuitCallCount);
                Assert.Contains(firstWorkbook, releasedObjects);
                Assert.DoesNotContain(cachedApplication, releasedObjects);

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession secondSession = strategy.OpenHiddenWorkbook(@"C:\Cases\cached-2.xlsx");
                Assert.Equal("app-cache", secondSession.RouteName);
                Assert.Same(cachedApplication, secondSession.Application);
                secondSession.Close();

                strategy.ShutdownHiddenApplicationCache();

                Assert.Equal(1, cachedApplication.QuitCallCount);
                Assert.Contains(cachedApplication, releasedObjects);
                Assert.Contains(logs, message => message.IndexOf("hidden-app-cache reused", StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        [Fact]
        public void OpenHiddenWorkbook_WithCacheEnabled_BypassesCacheWhileCachedApplicationIsInUse()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName, "1");
                Environment.SetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName, "60");
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application cachedApplication = CreateHiddenApplication();
                Excel.Application bypassApplication = CreateHiddenApplication();
                var strategy = CreateStrategy(logs, releasedObjects, new Queue<Excel.Application>(new[] { cachedApplication, bypassApplication }));

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession cachedSession = strategy.OpenHiddenWorkbook(@"C:\Cases\cached.xlsx");
                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession bypassSession = strategy.OpenHiddenWorkbook(@"C:\Cases\bypass.xlsx");

                Assert.Equal("app-cache", cachedSession.RouteName);
                Assert.Equal("app-cache-bypass-inuse", bypassSession.RouteName);
                Assert.Same(cachedApplication, cachedSession.Application);
                Assert.Same(bypassApplication, bypassSession.Application);

                bypassSession.Close();
                cachedSession.Close();
                strategy.ShutdownHiddenApplicationCache();

                Assert.Equal(1, bypassApplication.QuitCallCount);
                Assert.Equal(1, cachedApplication.QuitCallCount);
                Assert.Contains(logs, message => message.IndexOf("bypassed because in-use", StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        [Fact]
        public void OpenHiddenWorkbook_WithCacheEnabled_AbortPoisonsAndDiscardsCachedApplication()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName, "1");
                Environment.SetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName, "60");
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application firstApplication = CreateHiddenApplication();
                Excel.Application secondApplication = CreateHiddenApplication();
                var strategy = CreateStrategy(logs, releasedObjects, new Queue<Excel.Application>(new[] { firstApplication, secondApplication }));

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession firstSession = strategy.OpenHiddenWorkbook(@"C:\Cases\abort.xlsx");
                firstSession.Abort();

                Assert.Equal(1, firstApplication.QuitCallCount);
                Assert.Contains(firstApplication, releasedObjects);

                CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession secondSession = strategy.OpenHiddenWorkbook(@"C:\Cases\after-abort.xlsx");
                Assert.Same(secondApplication, secondSession.Application);
                secondSession.Close();
                strategy.ShutdownHiddenApplicationCache();

                Assert.Equal(1, secondApplication.QuitCallCount);
                Assert.Contains(logs, message => message.IndexOf("poisoned", StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        [Fact]
        public void OpenHiddenWorkbook_WithCacheEnabled_CleansUpCreatedApplication_WhenWorkbookOpenThrows()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName, "1");
                var logs = new List<string>();
                var releasedObjects = new List<object>();
                Excel.Application hiddenApplication = CreateHiddenApplication();
                hiddenApplication.Workbooks.OpenBehavior = (_, __, ___) => throw new InvalidOperationException("boom");
                var strategy = CreateStrategy(logs, releasedObjects, hiddenApplication);

                Assert.Throws<InvalidOperationException>(() => strategy.OpenHiddenWorkbook(@"C:\Cases\failure.xlsx"));

                Assert.Equal(1, hiddenApplication.QuitCallCount);
                Assert.Contains(hiddenApplication, releasedObjects);
                Assert.Contains(logs, message => message.IndexOf("poisoned", StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        [Fact]
        public void OpenHiddenForCaseDisplay_RestoresExcelState_OnSuccess()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                var logs = new List<string>();
                Excel.Application application = new Excel.Application
                {
                    ScreenUpdating = true,
                    EnableEvents = true,
                    DisplayAlerts = true,
                    ActiveWindow = new Excel.Window { Visible = true }
                };
                var strategy = new CaseWorkbookOpenStrategy(application, new WorkbookRoleResolver(), OrchestrationTestSupport.CreateLogger(logs));

                Excel.Workbook workbook = strategy.OpenHiddenForCaseDisplay(@"C:\Cases\display.xlsx");

                Assert.NotNull(workbook);
                Assert.True(application.ScreenUpdating);
                Assert.True(application.EnableEvents);
                Assert.True(application.DisplayAlerts);
            }
        }

        [Fact]
        public void OpenHiddenForCaseDisplay_RestoresExcelState_OnOpenFailure()
        {
            using (new HiddenRouteEnvironmentScope())
            {
                var logs = new List<string>();
                Excel.Window previousWindow = new Excel.Window { Visible = true };
                Excel.Application application = new Excel.Application
                {
                    ScreenUpdating = true,
                    EnableEvents = true,
                    DisplayAlerts = true,
                    ActiveWindow = previousWindow
                };
                application.Workbooks.OpenBehavior = (_, __, ___) => throw new InvalidOperationException("boom");
                var strategy = new CaseWorkbookOpenStrategy(application, new WorkbookRoleResolver(), OrchestrationTestSupport.CreateLogger(logs));

                Assert.Throws<InvalidOperationException>(() => strategy.OpenHiddenForCaseDisplay(@"C:\Cases\display-failure.xlsx"));

                Assert.True(application.ScreenUpdating);
                Assert.True(application.EnableEvents);
                Assert.True(application.DisplayAlerts);
                Assert.True(previousWindow.Visible);
                Assert.True(previousWindow.Activated);
            }
        }

        private static CaseWorkbookOpenStrategy CreateStrategy(List<string> logs, List<object> releasedObjects, params Excel.Application[] hiddenApplications)
        {
            return CreateStrategy(logs, releasedObjects, new Queue<Excel.Application>(hiddenApplications));
        }

        private static CaseWorkbookOpenStrategy CreateStrategy(List<string> logs, List<object> releasedObjects, Queue<Excel.Application> hiddenApplications)
        {
            return new CaseWorkbookOpenStrategy(
                new Excel.Application(),
                new WorkbookRoleResolver(),
                OrchestrationTestSupport.CreateLogger(logs),
                hiddenApplicationFactory: () =>
                {
                    if (hiddenApplications == null || hiddenApplications.Count == 0)
                    {
                        throw new InvalidOperationException("No hidden application prepared for test.");
                    }

                    return hiddenApplications.Dequeue();
                },
                releaseComObject: releasedObjects == null
                    ? null
                    : new Action<object>(comObject => releasedObjects.Add(comObject)));
        }

        private static Excel.Application CreateHiddenApplication()
        {
            var application = new Excel.Application
            {
                Ready = true,
                Visible = false,
                DisplayAlerts = false,
                ScreenUpdating = false,
                EnableEvents = false,
                UserControl = false
            };
            int windowSequence = 0;
            application.Workbooks.OpenBehavior = (filename, _, _) =>
            {
                var workbook = new Excel.Workbook
                {
                    FullName = filename ?? string.Empty,
                    Name = Path.GetFileName(filename ?? string.Empty),
                    Path = Path.GetDirectoryName(filename ?? string.Empty) ?? string.Empty
                };
                workbook.Windows.Add(new Excel.Window
                {
                    Visible = true,
                    Hwnd = ++windowSequence
                });
                return workbook;
            };
            return application;
        }

        private sealed class HiddenRouteEnvironmentScope : IDisposable
        {
            private readonly string _dedicatedHiddenInnerSave;
            private readonly string _legacyDedicatedHiddenInnerSaveAlias;
            private readonly string _hiddenApplicationCache;
            private readonly string _hiddenApplicationCacheIdleSeconds;

            internal HiddenRouteEnvironmentScope()
            {
                _dedicatedHiddenInnerSave = Environment.GetEnvironmentVariable(DedicatedHiddenInnerSaveEnvironmentVariableName);
                _legacyDedicatedHiddenInnerSaveAlias = Environment.GetEnvironmentVariable(LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName);
                _hiddenApplicationCache = Environment.GetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName);
                _hiddenApplicationCacheIdleSeconds = Environment.GetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName);
                Environment.SetEnvironmentVariable(DedicatedHiddenInnerSaveEnvironmentVariableName, null);
                Environment.SetEnvironmentVariable(LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName, null);
                Environment.SetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName, null);
                Environment.SetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName, null);
                Excel.Application.ResetCreatedApplications();
            }

            public void Dispose()
            {
                Environment.SetEnvironmentVariable(DedicatedHiddenInnerSaveEnvironmentVariableName, _dedicatedHiddenInnerSave);
                Environment.SetEnvironmentVariable(LegacyDedicatedHiddenInnerSaveAliasEnvironmentVariableName, _legacyDedicatedHiddenInnerSaveAlias);
                Environment.SetEnvironmentVariable(HiddenApplicationCacheEnvironmentVariableName, _hiddenApplicationCache);
                Environment.SetEnvironmentVariable(HiddenApplicationCacheIdleSecondsEnvironmentVariableName, _hiddenApplicationCacheIdleSeconds);
                Excel.Application.ResetCreatedApplications();
            }
        }
    }
}
