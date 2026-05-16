using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class AddInStartupBoundaryCoordinatorTests
    {
        [Fact]
        public void RunAfterApplicationEventsHooked_ShowsHomeBeforeStartupRefresh_WhenStartupPolicyAllows()
        {
            var calls = new List<string>();
            var coordinator = CreateCoordinator(
                new Excel.Application(),
                new List<string>(),
                shouldShowHome: true,
                clearHomeWorkbookBinding: reason => calls.Add("clear:" + reason),
                showKernelHomePlaceholder: () => calls.Add("show"),
                refreshTaskPane: (reason, workbook, window) => calls.Add("refresh:" + reason));

            coordinator.RunAfterApplicationEventsHooked();

            Assert.Equal(
                new[]
                {
                    "clear:ThisAddIn.TryShowKernelHomeFormOnStartup",
                    "show",
                    "refresh:Startup"
                },
                calls);
        }

        [Fact]
        public void ExecuteManagedCloseStartupGuard_QuitsEmptyStartupAndLeavesDisplayAlertsFalseOnSuccess()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true,
                Visible = false
            };
            var coordinator = CreateCoordinator(application, new List<string>());

            coordinator.ExecuteManagedCloseStartupGuard(CreateValidMarkerResult());

            Assert.Equal(1, application.QuitCallCount);
            Assert.False(application.DisplayAlerts);
        }

        [Fact]
        public void ExecuteManagedCloseStartupGuard_RestoresDisplayAlertsWhenQuitFails()
        {
            var logs = new List<string>();
            var application = new Excel.Application
            {
                DisplayAlerts = true,
                Visible = false,
                QuitBehavior = () => throw new InvalidOperationException("quit failed")
            };
            var coordinator = CreateCoordinator(application, logs);

            coordinator.ExecuteManagedCloseStartupGuard(CreateValidMarkerResult());

            Assert.Equal(1, application.QuitCallCount);
            Assert.True(application.DisplayAlerts);
            Assert.Contains(logs, message => message.Contains("Managed close startup guard quit failed."));
        }

        [Fact]
        public void ExecuteManagedCloseStartupGuard_SkipsQuit_WhenWorkbookOpenWasObserved()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true,
                Visible = false
            };
            var coordinator = CreateCoordinator(application, new List<string>());

            coordinator.MarkWorkbookOpenObserved();
            coordinator.ExecuteManagedCloseStartupGuard(CreateValidMarkerResult());

            Assert.Equal(0, application.QuitCallCount);
            Assert.True(application.DisplayAlerts);
        }

        [Fact]
        public void ExecuteManagedCloseStartupGuard_SkipsQuit_WhenVisibleNonKernelWorkbookExists()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true,
                Visible = false
            };
            application.Workbooks.Add(new Excel.Workbook
            {
                FullName = @"C:\Cases\customer.xlsx",
                Name = "customer.xlsx"
            });
            var coordinator = CreateCoordinator(application, new List<string>());

            coordinator.ExecuteManagedCloseStartupGuard(CreateValidMarkerResult());

            Assert.Equal(0, application.QuitCallCount);
            Assert.True(application.DisplayAlerts);
        }

        private static AddInStartupBoundaryCoordinator CreateCoordinator(
            Excel.Application application,
            List<string> logs,
            bool shouldShowHome = false,
            Action<string> clearHomeWorkbookBinding = null,
            Action showKernelHomePlaceholder = null,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPane = null)
        {
            return new AddInStartupBoundaryCoordinator(
                application,
                new Logger(logs.Add),
                null,
                () => shouldShowHome,
                () => "startup-state",
                clearHomeWorkbookBinding ?? (_ => { }),
                showKernelHomePlaceholder ?? (() => { }),
                refreshTaskPane ?? ((reason, workbook, window) => { }),
                () => application.ActiveWorkbook,
                workbook => workbook == null ? string.Empty : workbook.Name,
                workbook => false);
        }

        private static ManagedWorkbookCloseMarkerReadResult CreateValidMarkerResult()
        {
            var marker = new ManagedWorkbookCloseMarker(
                ManagedWorkbookCloseMarkerKind.CaseClose,
                DateTime.UtcNow,
                ManagedWorkbookCloseMarkerStore.DefaultTimeToLiveSeconds,
                @"C:\Cases\closed.xlsx");
            return ManagedWorkbookCloseMarkerReadResult.Valid(
                @"C:\logs\managed-workbook-close.marker",
                marker,
                TimeSpan.FromMilliseconds(10));
        }
    }
}
