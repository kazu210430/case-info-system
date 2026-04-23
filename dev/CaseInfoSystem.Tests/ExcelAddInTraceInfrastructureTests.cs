using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    [CollectionDefinition("ExcelAddInTraceInfrastructure", DisableParallelization = true)]
    public sealed class ExcelAddInTraceInfrastructureCollection
    {
    }

    [Collection("ExcelAddInTraceInfrastructure")]
    public class ExcelAddInTraceInfrastructureTests
    {
        private const string DetailedStartupDiagnosticsEnvironmentVariableName = "CASEINFOSYSTEM_EXCEL_STARTUP_TRACE";
        private const string FallbackDirectoryName = "CaseInfoSystem.ExcelAddIn";
        private const string TraceLogFileName = "CaseInfoSystem.ExcelAddIn_trace.log";

        [Fact]
        public void Trace_WhenDetailedStartupDiagnosticsAreDisabled_LogsBasicModeWithoutHeavySnapshotFields()
        {
            using (new StartupTraceEnvironmentScope())
            {
                Environment.SetEnvironmentVariable(DetailedStartupDiagnosticsEnvironmentVariableName, null);
                var logs = new List<string>();

                ExcelProcessLaunchContextTracer.Trace(OrchestrationTestSupport.CreateLogger(logs));

                string log = Assert.Single(logs);
                Assert.Contains("Process launch context.", log);
                Assert.Contains("startupDiagnosticsMode=basic", log);
                Assert.DoesNotContain(", parentPid=", log);
                Assert.DoesNotContain(", commandLine=", log);
                Assert.DoesNotContain("excelProcesses=[", log);
            }
        }

        [Fact]
        public void Write_WhenPrimaryWriteFails_WritesFallbackLog()
        {
            using (var scope = new TraceLogFileScope())
            {
                string message = "fallback-only-" + Guid.NewGuid().ToString("N");

                try
                {
                    string primaryLogDirectory = Path.GetDirectoryName(scope.PrimaryLogPath);
                    if (!string.IsNullOrEmpty(primaryLogDirectory))
                    {
                        Directory.CreateDirectory(primaryLogDirectory);
                    }

                    File.WriteAllText(scope.PrimaryLogPath, string.Empty);
                    using (new FileStream(scope.PrimaryLogPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        ExcelAddInTraceLogWriter.Write(message);
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    ExcelAddInTraceLogWriter.Write(message);
                }

                if (File.Exists(scope.PrimaryLogPath))
                {
                    Assert.DoesNotContain(message, File.ReadAllText(scope.PrimaryLogPath));
                }

                Assert.True(File.Exists(scope.FallbackLogPath));
                string fallbackLog = File.ReadAllText(scope.FallbackLogPath);
                Assert.Contains(message, fallbackLog);
                Assert.Contains("[primaryWriteFailed][tempFallback]", fallbackLog);
            }
        }

        [Fact]
        public void ResolvePrimarySystemRootPath_WhenLocalDocumentsContainsSystemRoot_PrefersLocalDocuments()
        {
            using (var scope = new SystemRootResolutionScope())
            {
                string localSystemRoot = scope.CreateSystemRoot(scope.LocalDocumentsPath);
                string redirectedSystemRoot = scope.CreateSystemRoot(scope.RedirectedDocumentsPath);

                string resolved = ExcelAddInTraceLogWriter.ResolvePrimarySystemRootPath(
                    scope.LocalDocumentsPath,
                    scope.RedirectedDocumentsPath);

                Assert.Equal(localSystemRoot, resolved);
                Assert.NotEqual(redirectedSystemRoot, resolved);
            }
        }

        [Fact]
        public void ResolvePrimarySystemRootPath_WhenOnlyRedirectedDocumentsContainsSystemRoot_UsesRedirectedDocuments()
        {
            using (var scope = new SystemRootResolutionScope())
            {
                string redirectedSystemRoot = scope.CreateSystemRoot(scope.RedirectedDocumentsPath);

                string resolved = ExcelAddInTraceLogWriter.ResolvePrimarySystemRootPath(
                    scope.LocalDocumentsPath,
                    scope.RedirectedDocumentsPath);

                Assert.Equal(redirectedSystemRoot, resolved);
            }
        }

        [Fact]
        public void ResolvePrimarySystemRootPath_WhenNoKnownSystemRootExists_FallsBackToRedirectedDocuments()
        {
            using (var scope = new SystemRootResolutionScope())
            {
                string resolved = ExcelAddInTraceLogWriter.ResolvePrimarySystemRootPath(
                    scope.LocalDocumentsPath,
                    scope.RedirectedDocumentsPath);

                Assert.Equal(Path.Combine(scope.RedirectedDocumentsPath, "案件情報System"), resolved);
            }
        }

        private sealed class StartupTraceEnvironmentScope : IDisposable
        {
            private readonly string _originalValue;

            internal StartupTraceEnvironmentScope()
            {
                _originalValue = Environment.GetEnvironmentVariable(DetailedStartupDiagnosticsEnvironmentVariableName);
            }

            public void Dispose()
            {
                Environment.SetEnvironmentVariable(DetailedStartupDiagnosticsEnvironmentVariableName, _originalValue);
            }
        }

        private sealed class TraceLogFileScope : IDisposable
        {
            private readonly FileSnapshot _primarySnapshot;
            private readonly FileSnapshot _fallbackSnapshot;

            internal TraceLogFileScope()
            {
                PrimaryLogPath = ExcelAddInTraceLogWriter.GetPrimaryTraceLogPath();
                FallbackLogPath = Path.Combine(Path.GetTempPath(), FallbackDirectoryName, TraceLogFileName);
                _primarySnapshot = FileSnapshot.Capture(PrimaryLogPath);
                _fallbackSnapshot = FileSnapshot.Capture(FallbackLogPath);
                EnsureFileAbsent(PrimaryLogPath);
                EnsureFileAbsent(FallbackLogPath);
            }

            internal string PrimaryLogPath { get; }

            internal string FallbackLogPath { get; }

            public void Dispose()
            {
                RestoreSnapshot(_primarySnapshot);
                RestoreSnapshot(_fallbackSnapshot);
            }

            private static void EnsureFileAbsent(string path)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
                catch
                {
                }
            }

            private static void RestoreSnapshot(FileSnapshot snapshot)
            {
                try
                {
                    if (snapshot.Existed)
                    {
                        string directoryPath = Path.GetDirectoryName(snapshot.Path);
                        if (!string.IsNullOrEmpty(directoryPath))
                        {
                            Directory.CreateDirectory(directoryPath);
                        }

                        File.WriteAllBytes(snapshot.Path, snapshot.Content ?? Array.Empty<byte>());
                        return;
                    }

                    if (File.Exists(snapshot.Path))
                    {
                        File.Delete(snapshot.Path);
                    }
                }
                catch
                {
                }
            }
        }

        private sealed class SystemRootResolutionScope : IDisposable
        {
            internal SystemRootResolutionScope()
            {
                RootPath = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.TraceRootResolution." + Guid.NewGuid().ToString("N"));
                LocalDocumentsPath = Path.Combine(RootPath, "LocalDocuments");
                RedirectedDocumentsPath = Path.Combine(RootPath, "RedirectedDocuments");
                Directory.CreateDirectory(LocalDocumentsPath);
                Directory.CreateDirectory(RedirectedDocumentsPath);
            }

            internal string RootPath { get; }

            internal string LocalDocumentsPath { get; }

            internal string RedirectedDocumentsPath { get; }

            internal string CreateSystemRoot(string documentsPath)
            {
                string systemRoot = Path.Combine(documentsPath, "案件情報System");
                Directory.CreateDirectory(Path.Combine(systemRoot, "Addins", "CaseInfoSystem.ExcelAddIn"));
                return systemRoot;
            }

            public void Dispose()
            {
                try
                {
                    if (Directory.Exists(RootPath))
                    {
                        Directory.Delete(RootPath, recursive: true);
                    }
                }
                catch
                {
                }
            }
        }

        private sealed class FileSnapshot
        {
            private FileSnapshot(string path, bool existed, byte[] content)
            {
                Path = path;
                Existed = existed;
                Content = content;
            }

            internal string Path { get; }

            internal bool Existed { get; }

            internal byte[] Content { get; }

            internal static FileSnapshot Capture(string path)
            {
                if (!File.Exists(path))
                {
                    return new FileSnapshot(path, false, null);
                }

                return new FileSnapshot(path, true, File.ReadAllBytes(path));
            }
        }
    }
}
