using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public sealed class KernelTemplateFolderOpenServiceTests : IDisposable
    {
        private readonly List<string> _temporaryDirectories = new List<string>();

        [Fact]
        public void TryOpen_WhenConfiguredTemplateDirectoryExists_UsesConfiguredDirectory()
        {
            using var harness = KernelTemplateFolderOpenHarness.Create(_temporaryDirectories);

            string configuredDirectory = harness.CreateDirectory("ConfiguredTemplates");
            harness.SetDocumentProperty("WORD_TEMPLATE_DIR", configuredDirectory);

            KernelTemplateFolderOpenService.OpenResult result = harness.Service.TryOpen(harness.KernelWorkbook);

            Assert.True(result.Success);
            Assert.Equal(configuredDirectory, harness.OpenedFolderPath);
        }

        [Fact]
        public void TryOpen_WhenConfiguredTemplateDirectoryIsUnavailable_FallsBackToSystemRootTemplateFolder()
        {
            using var harness = KernelTemplateFolderOpenHarness.Create(_temporaryDirectories);

            string fallbackDirectory = harness.CreateDirectory(Path.Combine("SystemRoot", "雛形"));
            harness.SetDocumentProperty("WORD_TEMPLATE_DIR", Path.Combine(harness.RootPath, "MissingTemplates"));

            KernelTemplateFolderOpenService.OpenResult result = harness.Service.TryOpen(harness.KernelWorkbook);

            Assert.True(result.Success);
            Assert.Equal(fallbackDirectory, harness.OpenedFolderPath);
        }

        [Fact]
        public void TryOpen_WhenResolvedFolderDoesNotExist_ReturnsFailureWithoutLaunchingExplorer()
        {
            using var harness = KernelTemplateFolderOpenHarness.Create(_temporaryDirectories);

            KernelTemplateFolderOpenService.OpenResult result = harness.Service.TryOpen(harness.KernelWorkbook);

            Assert.False(result.Success);
            Assert.Null(harness.OpenedFolderPath);
            Assert.Contains("雛形フォルダが見つかりませんでした。", result.FailureMessage);
        }

        [Fact]
        public void TryOpen_WhenSystemRootCannotBeResolved_ReturnsFailure()
        {
            using var harness = KernelTemplateFolderOpenHarness.Create(_temporaryDirectories, includeSystemRootProperty: false);
            harness.KernelWorkbook.Path = string.Empty;

            KernelTemplateFolderOpenService.OpenResult result = harness.Service.TryOpen(harness.KernelWorkbook);

            Assert.False(result.Success);
            Assert.Contains("雛形フォルダを解決できませんでした。", result.FailureMessage);
        }

        public void Dispose()
        {
            foreach (string temporaryDirectory in _temporaryDirectories)
            {
                try
                {
                    if (Directory.Exists(temporaryDirectory))
                    {
                        Directory.Delete(temporaryDirectory, recursive: true);
                    }
                }
                catch
                {
                }
            }
        }

        private sealed class KernelTemplateFolderOpenHarness : IDisposable
        {
            private KernelTemplateFolderOpenHarness(
                string rootPath,
                Excel.Workbook kernelWorkbook,
                KernelTemplateFolderOpenService service)
            {
                RootPath = rootPath;
                KernelWorkbook = kernelWorkbook;
                Service = service;
            }

            internal string RootPath { get; }

            internal Excel.Workbook KernelWorkbook { get; }

            internal KernelTemplateFolderOpenService Service { get; }

            internal string OpenedFolderPath { get; private set; }

            internal static KernelTemplateFolderOpenHarness Create(
                ICollection<string> temporaryDirectories,
                bool includeSystemRootProperty = true)
            {
                string rootPath = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.TemplateFolderOpen." + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(rootPath);
                temporaryDirectories.Add(rootPath);

                string systemRoot = Path.Combine(rootPath, "SystemRoot");
                Directory.CreateDirectory(systemRoot);

                var logger = new Logger(_ => { });
                var application = new Excel.Application();
                var pathCompatibilityService = new PathCompatibilityService(logger);
                var kernelWorkbook = new Excel.Workbook
                {
                    Name = "案件情報System_Kernel.xlsx",
                    FullName = Path.Combine(systemRoot, "案件情報System_Kernel.xlsx"),
                    Path = systemRoot,
                    CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                };
                application.Workbooks.Add(kernelWorkbook);

                if (includeSystemRootProperty)
                {
                    SetDocumentProperty(kernelWorkbook, "SYSTEM_ROOT", systemRoot);
                }

                var folderWindowService = new FolderWindowService(pathCompatibilityService, logger);
                var pathResolver = new KernelTemplateFolderPathResolver(
                    pathCompatibilityService,
                    TryGetDocumentProperty,
                    workbook => workbook == null ? string.Empty : workbook.Path ?? string.Empty);
                var harness = new KernelTemplateFolderOpenHarness(
                    rootPath,
                    kernelWorkbook,
                    new KernelTemplateFolderOpenService(pathResolver, pathCompatibilityService, folderWindowService, logger));

                folderWindowService.Hooks = new FolderWindowService.TestHooks
                {
                    StartFolderProcess = (folderPath, reason) =>
                    {
                        harness.OpenedFolderPath = folderPath;
                        return true;
                    }
                };

                return harness;
            }

            internal void SetDocumentProperty(string propertyName, string value)
            {
                SetDocumentProperty(KernelWorkbook, propertyName, value);
            }

            internal string CreateDirectory(string relativePath)
            {
                string fullPath = Path.Combine(RootPath, relativePath);
                Directory.CreateDirectory(fullPath);
                return fullPath;
            }

            private static void SetDocumentProperty(Excel.Workbook workbook, string propertyName, string value)
            {
                if (!(workbook?.CustomDocumentProperties is IDictionary<string, string> properties)
                    || string.IsNullOrWhiteSpace(propertyName))
                {
                    return;
                }

                properties[propertyName] = value ?? string.Empty;
            }

            private static string TryGetDocumentProperty(Excel.Workbook workbook, string propertyName)
            {
                if (!(workbook?.CustomDocumentProperties is IDictionary<string, string> properties)
                    || string.IsNullOrWhiteSpace(propertyName)
                    || !properties.TryGetValue(propertyName, out string value))
                {
                    return string.Empty;
                }

                return value ?? string.Empty;
            }

            public void Dispose()
            {
                try
                {
                    KernelWorkbook?.Close(false, null, null);
                }
                catch
                {
                }
            }
        }
    }
}
