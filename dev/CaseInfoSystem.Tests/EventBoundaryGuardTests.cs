using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class EventBoundaryGuardTests
    {
        [Fact]
        public void Execute_WhenHandlerThrows_SwallowsAndLogs()
        {
            var logs = new List<string>();

            var exception = Record.Exception(() =>
                EventBoundaryGuard.Execute(
                    OrchestrationTestSupport.CreateLogger(logs),
                    "Application_SheetChange",
                    () => throw new InvalidOperationException("boom")));

            Assert.Null(exception);
            Assert.Contains(logs, message => message.Contains("Application_SheetChange failed."));
        }

        [Fact]
        public void ExecuteCancelable_WhenHandlerThrows_SetsCancelAndLogs()
        {
            var logs = new List<string>();
            bool cancel = false;

            var exception = Record.Exception(() =>
                EventBoundaryGuard.ExecuteCancelable(
                    OrchestrationTestSupport.CreateLogger(logs),
                    "Application_WorkbookBeforeClose",
                    ref cancel,
                    () => throw new InvalidOperationException("boom")));

            Assert.Null(exception);
            Assert.True(cancel);
            Assert.Contains(logs, message => message.Contains("Application_WorkbookBeforeClose failed."));
        }
    }
}
