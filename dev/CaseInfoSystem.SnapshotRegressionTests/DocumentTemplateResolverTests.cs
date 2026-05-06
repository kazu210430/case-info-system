using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Core;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    public class DocumentTemplateResolverTests
    {
        [Fact]
        public void Resolve_WhenKeyIsBlankAfterTrim_ReturnsNullWithoutLookup()
        {
            using ResolverHarness harness = ResolverHarness.Create();

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, " \t ");

            Assert.Null(templateSpec);
            Assert.Equal(0, harness.LookupReader.CallCount);
        }

        [Fact]
        public void Resolve_WhenLookupMisses_ReturnsNull()
        {
            using ResolverHarness harness = ResolverHarness.Create(lookupSucceeds: false);

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, " 1 ");

            Assert.Null(templateSpec);
            Assert.Equal(1, harness.LookupReader.CallCount);
            Assert.Equal("1", harness.LookupReader.LastKey);
            Assert.Same(harness.Workbook, harness.LookupReader.LastWorkbook);
        }

        [Fact]
        public void Resolve_WhenLookupHit_UsesNormalizedCallerKeyAndBasicMetadata()
        {
            using ResolverHarness harness = ResolverHarness.Create(
                lookupResult: CreateLookupResult(
                    key: "99",
                    documentName: "委任状",
                    templateFileName: "01_委任状.docx",
                    resolutionSource: DocumentTemplateResolutionSource.SnapshotCache),
                wordTemplateDirectory: @"C:\Templates");

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, " 1 ");

            Assert.NotNull(templateSpec);
            Assert.Equal("1", templateSpec.Key);
            Assert.Equal("doc", templateSpec.ActionKind);
            Assert.Equal("委任状", templateSpec.DocumentName);
            Assert.Equal("01_委任状.docx", templateSpec.TemplateFileName);
            Assert.Equal(@"C:\Templates\01_委任状.docx", templateSpec.TemplatePath);
            Assert.Equal(DocumentTemplateResolutionSource.SnapshotCache, templateSpec.ResolutionSource);
            Assert.Equal("1", harness.LookupReader.LastKey);
        }

        [Fact]
        public void Resolve_WhenWordTemplateDirAndSystemRootBothExist_PrefersWordTemplateDir()
        {
            using ResolverHarness harness = ResolverHarness.Create(
                lookupResult: CreateLookupResult(templateFileName: "01_委任状.docx"),
                wordTemplateDirectory: @"C:\ConfiguredTemplates",
                systemRoot: @"C:\SystemRoot");

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal(@"C:\ConfiguredTemplates\01_委任状.docx", templateSpec.TemplatePath);
        }

        [Fact]
        public void Resolve_WhenWordTemplateDirMissing_UsesSystemRootTemplateFolder()
        {
            using ResolverHarness harness = ResolverHarness.Create(
                lookupResult: CreateLookupResult(templateFileName: "01_委任状.docx"),
                systemRoot: @"C:\SystemRoot");

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal(Path.Combine(@"C:\SystemRoot", "雛形", "01_委任状.docx"), templateSpec.TemplatePath);
        }

        [Fact]
        public void Resolve_WhenTemplateDirectoryCannotBeDerived_ReturnsSpecWithEmptyTemplatePath()
        {
            using ResolverHarness harness = ResolverHarness.Create(
                lookupResult: CreateLookupResult(
                    documentName: "委任状",
                    templateFileName: "01_委任状.docx",
                    resolutionSource: DocumentTemplateResolutionSource.SnapshotCache));

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal("1", templateSpec.Key);
            Assert.Equal("委任状", templateSpec.DocumentName);
            Assert.Equal("01_委任状.docx", templateSpec.TemplateFileName);
            Assert.Equal(string.Empty, templateSpec.TemplatePath);
            Assert.Equal("doc", templateSpec.ActionKind);
            Assert.Equal(DocumentTemplateResolutionSource.SnapshotCache, templateSpec.ResolutionSource);
        }

        [Fact]
        public void Resolve_WhenLookupReportsMasterCatalog_PreservesResolutionSource()
        {
            using ResolverHarness harness = ResolverHarness.Create(
                lookupResult: CreateLookupResult(
                    documentName: "委任状",
                    templateFileName: "01_委任状.docx",
                    resolutionSource: DocumentTemplateResolutionSource.MasterCatalog),
                systemRoot: @"C:\SystemRoot");

            DocumentTemplateSpec templateSpec = harness.Resolver.Resolve(harness.Workbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal(DocumentTemplateResolutionSource.MasterCatalog, templateSpec.ResolutionSource);
        }

        [Theory]
        [InlineData(@"C:\Templates\01_template.docx")]
        [InlineData(@"C:\Templates\01_template.dotx")]
        [InlineData(@"C:\Templates\01_template.dotm")]
        public void IsSupportedWordTemplate_WhenExtensionIsSupported_ReturnsTrue(string templatePath)
        {
            var templateSpec = new DocumentTemplateSpec
            {
                TemplatePath = templatePath
            };

            bool supported = DocumentTemplateResolver.IsSupportedWordTemplate(templateSpec);

            Assert.True(supported);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData(@"C:\Templates\01_template.doc")]
        [InlineData(@"C:\Templates\01_template.docm")]
        public void IsSupportedWordTemplate_WhenExtensionIsNotSupported_ReturnsFalse(string templatePath)
        {
            var templateSpec = new DocumentTemplateSpec
            {
                TemplatePath = templatePath
            };

            bool supported = DocumentTemplateResolver.IsSupportedWordTemplate(templateSpec);

            Assert.False(supported);
        }

        private static DocumentTemplateLookupResult CreateLookupResult(
            string key = "1",
            string documentName = "Document",
            string templateFileName = "01_template.docx",
            DocumentTemplateResolutionSource resolutionSource = DocumentTemplateResolutionSource.SnapshotCache)
        {
            return new DocumentTemplateLookupResult
            {
                Key = key ?? string.Empty,
                DocumentName = documentName ?? string.Empty,
                TemplateFileName = templateFileName ?? string.Empty,
                ResolutionSource = resolutionSource
            };
        }

        private sealed class ResolverHarness : IDisposable
        {
            private ResolverHarness(
                DocumentTemplateResolver resolver,
                StubDocumentTemplateLookupReader lookupReader,
                Excel.Workbook workbook)
            {
                Resolver = resolver;
                LookupReader = lookupReader;
                Workbook = workbook;
            }

            internal DocumentTemplateResolver Resolver { get; }

            internal StubDocumentTemplateLookupReader LookupReader { get; }

            internal Excel.Workbook Workbook { get; }

            internal static ResolverHarness Create(
                DocumentTemplateLookupResult lookupResult = null,
                bool lookupSucceeds = true,
                string wordTemplateDirectory = null,
                string systemRoot = null)
            {
                var logger = new Logger(_ => { });
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

                return new ResolverHarness(
                    new DocumentTemplateResolver(excelInteropService, pathCompatibilityService, lookupReader, logger),
                    lookupReader,
                    workbook);
            }

            public void Dispose()
            {
                Workbook?.Close(false, null, null);
            }
        }

        private sealed class StubDocumentTemplateLookupReader : IDocumentTemplateLookupReader
        {
            internal int CallCount { get; private set; }

            internal string LastKey { get; private set; }

            internal Excel.Workbook LastWorkbook { get; private set; }

            internal bool LookupSucceeds { get; set; }

            internal DocumentTemplateLookupResult Result { get; set; }

            public bool TryResolveWithMasterFallback(Excel.Workbook workbook, string key, out DocumentTemplateLookupResult result)
            {
                CallCount++;
                LastWorkbook = workbook;
                LastKey = key;
                result = LookupSucceeds ? Result : null;
                return LookupSucceeds && result != null;
            }
        }
    }
}
