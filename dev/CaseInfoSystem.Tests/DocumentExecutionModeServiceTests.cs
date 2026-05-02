using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class DocumentExecutionModeServiceTests
    {
        [Fact]
        public void GetMode_WhenSamePathTimestampChanges_ReloadsMode()
        {
            DateTime initialTimestamp = new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc);
            DateTime updatedTimestamp = initialTimestamp.AddMinutes(5);
            DateTime currentTimestamp = initialTimestamp;
            string currentValue = "PilotOnly";

            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => currentTimestamp,
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { currentValue }
                });

            Assert.Equal(DocumentExecutionMode.PilotOnly, service.GetMode());

            currentTimestamp = updatedTimestamp;
            currentValue = "Disabled";

            Assert.Equal(DocumentExecutionMode.Disabled, service.GetMode());
        }

        [Fact]
        public void GetMode_WhenPathChanges_ReloadsMode()
        {
            string currentPath = @"C:\runtime-a\DocumentExecutionMode.txt";
            DateTime currentTimestamp = new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc);
            var valuesByPath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [@"C:\runtime-a\DocumentExecutionMode.txt"] = "PilotOnly",
                [@"C:\runtime-b\DocumentExecutionMode.txt"] = "Disabled"
            };

            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => currentTimestamp,
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { valuesByPath[path] },
                    ResolveModeFilePath = () => currentPath
                });

            Assert.Equal(DocumentExecutionMode.PilotOnly, service.GetMode());

            currentPath = @"C:\runtime-b\DocumentExecutionMode.txt";

            Assert.Equal(DocumentExecutionMode.Disabled, service.GetMode());
        }

        [Fact]
        public void GetMode_WhenReloadFails_KeepsPreviousValidMode()
        {
            string modePath = @"C:\runtime\DocumentExecutionMode.txt";
            DateTime initialTimestamp = new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc);
            DateTime failingTimestamp = initialTimestamp.AddMinutes(5);
            DateTime currentTimestamp = initialTimestamp;
            bool throwOnRead = false;

            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => currentTimestamp,
                    ModeFileExists = path => true,
                    ReadModeFileLines = path =>
                    {
                        if (throwOnRead)
                        {
                            throw new InvalidOperationException("read failed");
                        }

                        return new[] { "PilotOnly" };
                    },
                    ResolveModeFilePath = () => modePath
                });

            Assert.Equal(DocumentExecutionMode.PilotOnly, service.GetMode());

            currentTimestamp = failingTimestamp;
            throwOnRead = true;

            Assert.Equal(DocumentExecutionMode.PilotOnly, service.GetMode());
        }

        [Fact]
        public void CanAttemptVstoExecution_WhenModeIsPilotOnly_ReturnsTrue()
        {
            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc),
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { "PilotOnly" },
                    ResolveModeFilePath = () => @"C:\runtime\DocumentExecutionMode.txt"
                });

            Assert.True(service.CanAttemptVstoExecution());
        }

        [Fact]
        public void GetMode_WhenValueIsUnknown_FallsBackToDisabled()
        {
            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc),
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { "UnexpectedMode" },
                    ResolveModeFilePath = () => @"C:\runtime\DocumentExecutionMode.txt"
                });

            Assert.Equal(DocumentExecutionMode.Disabled, service.GetMode());
        }
    }
}
