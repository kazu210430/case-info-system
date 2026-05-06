using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Core;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    public class DocumentExecutionEligibilityServiceTests
    {
        [Fact]
        public void Evaluate_WhenTemplateExtensionIsDoc_ReturnsUnsupported()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            string expectedPath = Path.Combine(templateDirectory, "01_template.doc");
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.doc"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("template type is not supported: " + expectedPath, result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenTemplateExtensionIsDocm_ReturnsUnsupported()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            string expectedPath = Path.Combine(templateDirectory, "01_template.docm");
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docm"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("template type is not supported: " + expectedPath, result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenTemplateExtensionIsDotm_ReturnsMacroEnabledRejection()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            string expectedPath = Path.Combine(templateDirectory, "01_template.dotm");
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.dotm"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("macro-enabled word template is routed to VBA: " + expectedPath, result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenLookupMisses_ReturnsFailClosed()
        {
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupSucceeds: false,
                outputFolder: @"C:\unused",
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("template spec was not resolved", result.Reason);
            Assert.Null(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenTemplateDirectoryCannotBeResolved_ReturnsFailClosed()
        {
            using var workspace = TestWorkspace.Create();
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docx"),
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("template type is not supported: ", result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenTemplateFileIsMissing_ReturnsFailClosed()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            string expectedPath = Path.Combine(templateDirectory, "01_template.docx");
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docx"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("template file was not found: " + expectedPath, result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenOutputFolderCannotBeResolved_ReturnsFailClosed()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            workspace.CreateFile(Path.Combine("Templates", "01_template.docx"));
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docx"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: string.Empty,
                caseContext: CreateCaseContextWithValues());

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("output folder could not be resolved", result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenCaseContextIsNull_ReturnsFailClosed()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            workspace.CreateFile(Path.Combine("Templates", "01_template.docx"));
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docx"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: null);

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("case context could not be resolved", result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenCaseSnapshotIsEmpty_ReturnsFailClosed()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            workspace.CreateFile(Path.Combine("Templates", "01_template.docx"));
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docx"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: new CaseContext
                {
                    CaseValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                });

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.False(result.CanExecuteInVsto);
            Assert.Equal("case snapshot could not be resolved", result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Null(result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenAllPreconditionsAreMet_ReturnsEligible()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            workspace.CreateFile(Path.Combine("Templates", "01_template.docx"));
            CaseContext caseContext = CreateCaseContextWithValues();
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult("Document", "01_template.docx"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: caseContext);

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.True(result.CanExecuteInVsto);
            Assert.Equal("eligible", result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Same(caseContext, result.CaseContext);
        }

        [Fact]
        public void Evaluate_WhenDocumentNameIsEmpty_ReturnsEligibleAndWarns()
        {
            using var workspace = TestWorkspace.Create();
            string templateDirectory = workspace.CreateDirectory("Templates");
            workspace.CreateFile(Path.Combine("Templates", "01_template.docx"));
            CaseContext caseContext = CreateCaseContextWithValues();
            using EligibilityHarness harness = EligibilityHarness.Create(
                lookupResult: CreateLookupResult(string.Empty, "01_template.docx"),
                wordTemplateDirectory: templateDirectory,
                outputFolder: workspace.CreateDirectory("Output"),
                caseContext: caseContext);

            DocumentExecutionEligibility result = harness.Service.Evaluate(harness.Workbook, "doc", "1");

            Assert.True(result.CanExecuteInVsto);
            Assert.Equal("eligible", result.Reason);
            Assert.NotNull(result.TemplateSpec);
            Assert.Equal(string.Empty, result.TemplateSpec.DocumentName);
            Assert.Same(caseContext, result.CaseContext);
            Assert.Contains(harness.Logs, message => message.Contains("DocumentExecutionEligibilityService found empty document name."));
        }

        private static DocumentTemplateLookupResult CreateLookupResult(string documentName, string templateFileName)
        {
            return new DocumentTemplateLookupResult
            {
                Key = "1",
                DocumentName = documentName ?? string.Empty,
                TemplateFileName = templateFileName ?? string.Empty,
                ResolutionSource = DocumentTemplateResolutionSource.SnapshotCache
            };
        }

        private static CaseContext CreateCaseContextWithValues()
        {
            return new CaseContext
            {
                CaseValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Customer_Name"] = "Customer"
                }
            };
        }

        private sealed class EligibilityHarness : IDisposable
        {
            private EligibilityHarness(
                DocumentExecutionEligibilityService service,
                Excel.Workbook workbook,
                List<string> logs)
            {
                Service = service;
                Workbook = workbook;
                Logs = logs;
            }

            internal DocumentExecutionEligibilityService Service { get; }

            internal Excel.Workbook Workbook { get; }

            internal List<string> Logs { get; }

            internal static EligibilityHarness Create(
                DocumentTemplateLookupResult lookupResult = null,
                bool lookupSucceeds = true,
                string wordTemplateDirectory = null,
                string systemRoot = null,
                string outputFolder = null,
                CaseContext caseContext = null)
            {
                var logs = new List<string>();
                var logger = new Logger(message => logs.Add(message ?? string.Empty));
                var application = new Excel.Application();
                var pathCompatibilityService = new PathCompatibilityService();
                var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);
                var workbook = new Excel.Workbook
                {
                    Name = "Case.xlsx",
                    FullName = @"C:\Cases\Case.xlsx",
                    Path = @"C:\Cases",
                    CustomDocumentProperties = new DocumentProperties()
                };
                application.Workbooks.Add(workbook);

                if (wordTemplateDirectory != null)
                {
                    excelInteropService.SetDocumentProperty(workbook, "WORD_TEMPLATE_DIR", wordTemplateDirectory);
                }

                if (systemRoot != null)
                {
                    excelInteropService.SetDocumentProperty(workbook, "SYSTEM_ROOT", systemRoot);
                }

                var lookupReader = new StubDocumentTemplateLookupReader
                {
                    LookupSucceeds = lookupSucceeds,
                    Result = lookupResult
                };
                var resolver = new DocumentTemplateResolver(excelInteropService, pathCompatibilityService, lookupReader, logger);
                var outputService = new DocumentOutputService
                {
                    OnResolveWorkbookFolder = _ => outputFolder
                };
                var caseContextFactory = new CaseContextFactory
                {
                    OnCreateForDocumentCreate = _ => caseContext
                };

                return new EligibilityHarness(
                    new DocumentExecutionEligibilityService(resolver, caseContextFactory, outputService, logger),
                    workbook,
                    logs);
            }

            public void Dispose()
            {
                Workbook?.Close(false, null, null);
            }
        }

        private sealed class StubDocumentTemplateLookupReader : IDocumentTemplateLookupReader
        {
            internal bool LookupSucceeds { get; set; }

            internal DocumentTemplateLookupResult Result { get; set; }

            public bool TryResolveWithMasterFallback(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result)
            {
                result = LookupSucceeds ? Result : null;
                return LookupSucceeds && result != null;
            }
        }

        private sealed class TestWorkspace : IDisposable
        {
            private TestWorkspace(string rootPath)
            {
                RootPath = rootPath;
            }

            internal string RootPath { get; }

            internal static TestWorkspace Create()
            {
                string rootPath = Path.Combine(
                    Path.GetTempPath(),
                    "CaseInfoSystem.DocumentExecutionEligibilityServiceTests",
                    Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(rootPath);
                return new TestWorkspace(rootPath);
            }

            internal string CreateDirectory(string relativePath)
            {
                string fullPath = Path.Combine(RootPath, relativePath ?? string.Empty);
                Directory.CreateDirectory(fullPath);
                return fullPath;
            }

            internal string CreateFile(string relativePath)
            {
                string fullPath = Path.Combine(RootPath, relativePath ?? string.Empty);
                string parentDirectory = Path.GetDirectoryName(fullPath) ?? string.Empty;
                if (parentDirectory.Length > 0)
                {
                    Directory.CreateDirectory(parentDirectory);
                }

                if (!File.Exists(fullPath))
                {
                    File.WriteAllText(fullPath, string.Empty);
                }

                return fullPath;
            }

            public void Dispose()
            {
                if (Directory.Exists(RootPath))
                {
                    Directory.Delete(RootPath, true);
                }
            }
        }
    }
}

namespace CaseInfoSystem.ExcelAddIn.App
{
    using System;
    using CaseInfoSystem.ExcelAddIn.Domain;
    using Excel = Microsoft.Office.Interop.Excel;

    internal sealed class CaseContextFactory
    {
        internal Func<Excel.Workbook, CaseContext> OnCreateForDocumentCreate { get; set; }

        internal CaseContext CreateForDocumentCreate(Excel.Workbook caseWorkbook)
        {
            return OnCreateForDocumentCreate == null ? null : OnCreateForDocumentCreate(caseWorkbook);
        }
    }

    internal sealed class DocumentOutputService
    {
        internal Func<Excel.Workbook, string> OnResolveWorkbookFolder { get; set; }

        internal string ResolveWorkbookFolder(Excel.Workbook workbook)
        {
            return OnResolveWorkbookFolder == null ? string.Empty : OnResolveWorkbookFolder(workbook) ?? string.Empty;
        }
    }
}
