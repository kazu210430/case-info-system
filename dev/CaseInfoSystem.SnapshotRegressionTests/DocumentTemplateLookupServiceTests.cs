using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    public class DocumentTemplateLookupServiceTests
    {
        [Fact]
        public void Resolve_WhenCaseCacheExists_ResolverAndPromptUseSameCaptionAndFile()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            DocumentNamePromptForm.OnTryPrompt = null;

            TestServices services = CreateServices(scenario.Application);
            services.ExcelInteropService.SetDocumentProperty(scenario.CaseWorkbook, "WORD_TEMPLATE_DIR", @"C:\Templates");
            scenario.Application.ActiveWorkbook = scenario.CaseWorkbook;
            scenario.Application.ActiveWindow = scenario.CaseWorkbook.Windows[1];

            DocumentTemplateSpec templateSpec = services.Resolver.Resolve(scenario.CaseWorkbook, "1");
            string promptedInitialName = null;
            DocumentNamePromptForm.OnTryPrompt = (owner, initialDocumentName) =>
            {
                promptedInitialName = initialDocumentName;
                return new DocumentNamePromptForm.PromptResult
                {
                    Accepted = false,
                    DocumentName = initialDocumentName
                };
            };

            bool accepted = services.PromptService.TryPrepare(scenario.CaseWorkbook, "1", out DocumentNameOverrideScope scope);

            Assert.False(accepted);
            Assert.Null(scope);
            Assert.NotNull(templateSpec);
            Assert.Equal("委任状", templateSpec.DocumentName);
            Assert.Equal("01_委任状.docx", templateSpec.TemplateFileName);
            Assert.Equal(@"C:\Templates\01_委任状.docx", templateSpec.TemplatePath);
            Assert.Equal(DocumentTemplateResolutionSource.SnapshotCache, templateSpec.ResolutionSource);
            Assert.Equal(templateSpec.DocumentName, promptedInitialName);
        }

        [Fact]
        public void Resolve_WhenCaseCacheMissing_FallsBackToMasterCatalogButPromptRemainsCacheOnly()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            DocumentNamePromptForm.OnTryPrompt = null;
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);

            TestServices services = CreateServices(scenario.Application);
            services.ExcelInteropService.SetDocumentProperty(scenario.CaseWorkbook, "WORD_TEMPLATE_DIR", @"C:\Templates");
            scenario.Application.ActiveWorkbook = scenario.CaseWorkbook;
            scenario.Application.ActiveWindow = scenario.CaseWorkbook.Windows[1];

            DocumentTemplateSpec templateSpec = services.Resolver.Resolve(scenario.CaseWorkbook, "1");
            string promptedInitialName = null;
            DocumentNamePromptForm.OnTryPrompt = (owner, initialDocumentName) =>
            {
                promptedInitialName = initialDocumentName;
                return new DocumentNamePromptForm.PromptResult
                {
                    Accepted = false,
                    DocumentName = initialDocumentName
                };
            };

            bool accepted = services.PromptService.TryPrepare(scenario.CaseWorkbook, "1", out DocumentNameOverrideScope scope);

            Assert.False(accepted);
            Assert.Null(scope);
            Assert.NotNull(templateSpec);
            Assert.Equal("委任状", templateSpec.DocumentName);
            Assert.Equal("01_委任状.docx", templateSpec.TemplateFileName);
            Assert.Equal(DocumentTemplateResolutionSource.MasterCatalog, templateSpec.ResolutionSource);
            Assert.Equal(string.Empty, promptedInitialName);
        }

        [Fact]
        public void Resolve_WhenKeyIsMissing_ReturnsNullAndPromptStartsBlank()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            DocumentNamePromptForm.OnTryPrompt = null;
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);

            TestServices services = CreateServices(scenario.Application);
            string promptedInitialName = null;
            DocumentNamePromptForm.OnTryPrompt = (owner, initialDocumentName) =>
            {
                promptedInitialName = initialDocumentName;
                return new DocumentNamePromptForm.PromptResult
                {
                    Accepted = false,
                    DocumentName = initialDocumentName
                };
            };

            DocumentTemplateSpec templateSpec = services.Resolver.Resolve(scenario.CaseWorkbook, "99");
            bool accepted = services.PromptService.TryPrepare(scenario.CaseWorkbook, "99", out DocumentNameOverrideScope scope);

            Assert.Null(templateSpec);
            Assert.False(accepted);
            Assert.Null(scope);
            Assert.Equal(string.Empty, promptedInitialName);
        }

        [Fact]
        public void CaseCacheReader_WhenCaseCacheMissing_DoesNotFallbackToMasterCatalog()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);

            TestServices services = CreateServices(scenario.Application);

            bool resolved = services.CaseCacheReader.TryResolveFromCaseCache(scenario.CaseWorkbook, "1", out DocumentTemplateLookupResult lookupResult);

            Assert.False(resolved);
            Assert.Null(lookupResult);
        }

        [Fact]
        public void Resolve_WhenWordTemplateDirectoryMissing_UsesSystemRootTemplateFolderForTemplatePath()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);

            TestServices services = CreateServices(scenario.Application);
            SnapshotBuilderScenario.InputRow row = CreateRows()[0];

            DocumentTemplateSpec templateSpec = services.Resolver.Resolve(scenario.CaseWorkbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal(row.TemplateFileName, templateSpec.TemplateFileName);
            Assert.Equal(
                Path.Combine(@"C:\SnapshotRegression\SystemRoot", "\u96DB\u5F62", row.TemplateFileName),
                templateSpec.TemplatePath);
            Assert.Equal(DocumentTemplateResolutionSource.SnapshotCache, templateSpec.ResolutionSource);
        }

        private static TestServices CreateServices(Excel.Application application)
        {
            var logger = new Logger(_ => { });
            var pathCompatibilityService = new PathCompatibilityService();
            var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);
            var taskPaneSnapshotCacheService = new TaskPaneSnapshotCacheService(excelInteropService, logger);
            var masterTemplateSheetReader = new MasterTemplateSheetReaderAdapter();
            var masterTemplateCatalogService = new MasterTemplateCatalogService(application, excelInteropService, pathCompatibilityService, masterTemplateSheetReader, logger);
            var lookupService = new DocumentTemplateLookupService(taskPaneSnapshotCacheService, masterTemplateCatalogService);

            return new TestServices(
                excelInteropService,
                lookupService,
                new DocumentTemplateResolver(excelInteropService, pathCompatibilityService, lookupService, logger),
                new DocumentNamePromptService(excelInteropService, lookupService, logger));
        }

        private static SnapshotBuilderScenario.InputRow[] CreateRows()
        {
            return new[]
            {
                new SnapshotBuilderScenario.InputRow
                {
                    Key = "1",
                    TemplateFileName = "01_委任状.docx",
                    Caption = "委任状",
                    TabName = "申請手続",
                    FillColor = 1111,
                    TabBackColor = 2222
                }
            };
        }

        private static void EnsureMasterWorkbookFileExists(string workbookFullName)
        {
            string path = workbookFullName ?? string.Empty;
            if (path.Length == 0)
            {
                return;
            }

            string directory = Path.GetDirectoryName(path) ?? string.Empty;
            if (directory.Length > 0)
            {
                Directory.CreateDirectory(directory);
            }

            if (!File.Exists(path))
            {
                File.WriteAllText(path, string.Empty);
            }
        }

        private sealed class TestServices
        {
            internal TestServices(
                ExcelInteropService excelInteropService,
                ICaseCacheDocumentTemplateReader caseCacheReader,
                DocumentTemplateResolver resolver,
                DocumentNamePromptService promptService)
            {
                ExcelInteropService = excelInteropService;
                CaseCacheReader = caseCacheReader;
                Resolver = resolver;
                PromptService = promptService;
            }

            internal ExcelInteropService ExcelInteropService { get; }

            internal ICaseCacheDocumentTemplateReader CaseCacheReader { get; }

            internal DocumentTemplateResolver Resolver { get; }

            internal DocumentNamePromptService PromptService { get; }
        }
    }
}
