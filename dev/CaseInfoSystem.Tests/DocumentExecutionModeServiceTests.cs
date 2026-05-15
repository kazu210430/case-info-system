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
            string currentValue = "WarmupEnabledProfileA";

            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => currentTimestamp,
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { currentValue }
                });

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileA, service.GetConfiguredMode());

            currentTimestamp = updatedTimestamp;
            currentValue = "WarmupEnabledProfileB";

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileB, service.GetConfiguredMode());
        }

        [Fact]
        public void GetMode_WhenPathChanges_ReloadsMode()
        {
            string currentPath = @"C:\runtime-a\DocumentExecutionMode.txt";
            DateTime currentTimestamp = new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc);
            var valuesByPath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [@"C:\runtime-a\DocumentExecutionMode.txt"] = "PilotOnly",
                [@"C:\runtime-b\DocumentExecutionMode.txt"] = "WarmupEnabledProfileB"
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

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileA, service.GetConfiguredMode());

            currentPath = @"C:\runtime-b\DocumentExecutionMode.txt";

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileB, service.GetConfiguredMode());
        }

        [Fact]
        public void GetMode_WhenPathChangesToDisabled_ReloadsDisabledMode()
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

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileA, service.GetConfiguredMode());

            currentPath = @"C:\runtime-b\DocumentExecutionMode.txt";

            Assert.Equal(DocumentExecutionMode.Disabled, service.GetConfiguredMode());
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

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileA, service.GetConfiguredMode());

            currentTimestamp = failingTimestamp;
            throwOnRead = true;

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileA, service.GetConfiguredMode());
        }

        [Fact]
        public void IsWordWarmupEnabled_WhenModeIsAllowlistedOnlyLegacyAlias_ReturnsTrueForCompatibility()
        {
            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc),
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { "AllowlistedOnly" },
                    ResolveModeFilePath = () => @"C:\runtime\DocumentExecutionMode.txt"
                });

            Assert.Equal(DocumentExecutionMode.WarmupEnabledProfileB, service.GetConfiguredMode());
            Assert.True(service.IsWordWarmupEnabled());
        }

        [Fact]
        public void IsWordWarmupEnabled_WhenModeIsDisabled_ReturnsFalse()
        {
            var service = new DocumentExecutionModeService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new ExcelInteropService(),
                new DocumentExecutionModeService.DocumentExecutionModeServiceTestHooks
                {
                    GetModeFileLastWriteTimeUtc = path => new DateTime(2026, 4, 18, 10, 0, 0, DateTimeKind.Utc),
                    ModeFileExists = path => true,
                    ReadModeFileLines = path => new[] { "Disabled" },
                    ResolveModeFilePath = () => @"C:\runtime\DocumentExecutionMode.txt"
                });

            Assert.False(service.IsWordWarmupEnabled());
        }
    }
}
