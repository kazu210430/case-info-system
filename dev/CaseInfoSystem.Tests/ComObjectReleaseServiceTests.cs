using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ComObjectReleaseServiceTests
    {
        [Fact]
        public void FinalRelease_WhenReleaseThrows_WritesWarnLogAndDoesNotThrow()
        {
            var logs = new List<string>();
            var testHooks = new ComObjectReleaseService.ComObjectReleaseServiceTestHooks
            {
                IsComObject = _ => true,
                FinalReleaseComObject = _ => throw new InvalidOperationException("release boom"),
                DebugWriteLine = _ => { }
            };
            var fakeComObject = new FakeComObject();

            Exception exception = Record.Exception(() =>
                ComObjectReleaseService.FinalRelease(
                    fakeComObject,
                    OrchestrationTestSupport.CreateLogger(logs),
                    "KernelUserDataReflectionService.CloseWorkbookQuietly target=Accounting",
                    testHooks));

            Assert.Null(exception);
            Assert.Contains(
                logs,
                message => message.Contains("WARN: COM cleanup release failed.", StringComparison.Ordinal)
                    && message.Contains("operation=FinalRelease", StringComparison.Ordinal)
                    && message.Contains(nameof(FakeComObject), StringComparison.Ordinal)
                    && message.Contains("context=KernelUserDataReflectionService.CloseWorkbookQuietly target=Accounting", StringComparison.Ordinal)
                    && message.Contains("exceptionType=InvalidOperationException", StringComparison.Ordinal)
                    && message.Contains("message=release boom", StringComparison.Ordinal));
        }

        [Fact]
        public void FinalRelease_WhenWarnLoggerThrows_DoesNotThrow()
        {
            var testHooks = new ComObjectReleaseService.ComObjectReleaseServiceTestHooks
            {
                IsComObject = _ => true,
                FinalReleaseComObject = _ => throw new InvalidOperationException("release boom"),
                DebugWriteLine = _ => { }
            };
            var fakeComObject = new FakeComObject();
            var logger = new Logger(_ => throw new InvalidOperationException("logger boom"));

            Exception exception = Record.Exception(() =>
                ComObjectReleaseService.FinalRelease(
                    fakeComObject,
                    logger,
                    "CaseWorkbookOpenStrategy.ReleaseComObject",
                    testHooks));

            Assert.Null(exception);
        }

        private sealed class FakeComObject
        {
        }
    }
}
