using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class AddInExecutionBoundaryCoordinatorTests
    {
        [Fact]
        public void Execute_RestoresScreenUpdatingAfterAction()
        {
            var application = new Excel.Application
            {
                ScreenUpdating = true
            };
            var coordinator = CreateCoordinator(application, new List<string>());

            bool observedSuspended = false;

            coordinator.Execute(() => observedSuspended = !application.ScreenUpdating);

            Assert.True(observedSuspended);
            Assert.True(application.ScreenUpdating);
        }

        [Fact]
        public void Execute_RestoresScreenUpdatingWhenActionThrows()
        {
            var application = new Excel.Application
            {
                ScreenUpdating = true
            };
            var coordinator = CreateCoordinator(application, new List<string>());

            Assert.Throws<InvalidOperationException>(
                () => coordinator.Execute(() => throw new InvalidOperationException("boom")));

            Assert.True(application.ScreenUpdating);
        }

        [Fact]
        public void Enter_IncrementsSuppressionCountAndLogsBoundary()
        {
            var logs = new List<string>();
            var coordinator = CreateCoordinator(new Excel.Application(), logs);

            using (coordinator.Enter("unit-test"))
            {
                Assert.Equal(1, coordinator.TaskPaneRefreshSuppressionCount);
            }

            Assert.Equal(0, coordinator.TaskPaneRefreshSuppressionCount);
            Assert.Contains(logs, message => message.Contains("action=suppress-enter")
                && message.Contains("source=ThisAddIn")
                && message.Contains("suppressionCount=1"));
            Assert.Contains(logs, message => message.Contains("action=suppress-exit")
                && message.Contains("source=ThisAddIn")
                && message.Contains("suppressionCount=0"));
        }

        private static AddInExecutionBoundaryCoordinator CreateCoordinator(
            Excel.Application application,
            List<string> logs)
        {
            return new AddInExecutionBoundaryCoordinator(
                application,
                new Logger(logs.Add),
                () => "active-state");
        }
    }
}
