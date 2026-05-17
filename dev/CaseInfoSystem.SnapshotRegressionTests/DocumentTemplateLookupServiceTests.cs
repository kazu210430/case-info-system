using System.Collections.Generic;
using System.Globalization;
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
        public void PromptAccepted_DoesNotLogDocumentNameValue()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            var logs = new List<string>();
            DocumentNameOverrideScope scope = null;
            DocumentNamePromptForm.OnTryPrompt = null;

            try
            {
                TestServices services = CreateServices(scenario.Application, logs);
                scenario.Application.ActiveWorkbook = scenario.CaseWorkbook;
                scenario.Application.ActiveWindow = scenario.CaseWorkbook.Windows[1];
                DocumentNamePromptForm.OnTryPrompt = (owner, initialDocumentName) => new DocumentNamePromptForm.PromptResult
                {
                    Accepted = true,
                    DocumentName = "山田太郎_確定文書名"
                };

                bool accepted = services.PromptService.TryPrepare(scenario.CaseWorkbook, "1", out scope);

                Assert.True(accepted);
                Assert.NotNull(scope);
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "山田太郎_確定文書名"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "委任状"));
                Assert.Contains(logs, message =>
                    ContainsOrdinal(message, "promptResult=Accepted")
                    && ContainsOrdinal(message, "finalDocumentNameProvided=True")
                    && ContainsOrdinal(message, "finalDocumentNameLength="));
            }
            finally
            {
                scope?.Dispose();
                DocumentNamePromptForm.OnTryPrompt = null;
            }
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
        public void Resolve_WhenMasterWorkbookAlreadyOpen_DoesNotChangeWindowVisibility()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);

            TestServices services = CreateServices(scenario.Application);
            scenario.MasterWorkbook.Windows[1].Visible = true;

            DocumentTemplateSpec templateSpec = services.Resolver.Resolve(scenario.CaseWorkbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal(DocumentTemplateResolutionSource.MasterCatalog, templateSpec.ResolutionSource);
            Assert.True(scenario.MasterWorkbook.Windows[1].Visible);
        }

        [Fact]
        public void Resolve_WhenMasterWorkbookOpenedForRead_HidesOnlyOpenedWorkbook()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);
            scenario.Application.Workbooks.Remove(scenario.MasterWorkbook);
            scenario.MasterWorkbook.Windows[1].Visible = true;
            scenario.Application.Workbooks.OpenBehavior = (_, __, ___) => scenario.MasterWorkbook;

            TestServices services = CreateServices(scenario.Application);

            DocumentTemplateSpec templateSpec = services.Resolver.Resolve(scenario.CaseWorkbook, "1");

            Assert.NotNull(templateSpec);
            Assert.Equal(DocumentTemplateResolutionSource.MasterCatalog, templateSpec.ResolutionSource);
            Assert.False(scenario.MasterWorkbook.Windows[1].Visible);
            Assert.DoesNotContain(scenario.MasterWorkbook, scenario.Application.Workbooks);
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
        public void PromotedCaseCacheReader_WhenCaseCacheMissing_DoesNotFallbackToMasterCatalog()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);

            TestServices services = CreateServices(scenario.Application);

            bool resolved = services.CaseCacheReader.TryEnsurePromotedCaseCacheThenResolve(scenario.CaseWorkbook, "1", out DocumentTemplateLookupResult lookupResult);

            Assert.False(resolved);
            Assert.Null(lookupResult);
        }

        [Fact]
        public void PromotedCaseCacheReader_WhenBaseSnapshotExists_PromotesCaseCacheAndUpdatesDocPropertiesBeforeResolve()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            string embeddedSnapshot = SnapshotLegacySerializer.Serialize(
                scenario.CaseWorkbook,
                42,
                scenario.Values,
                scenario.FillColors,
                scenario.TabBackColors);

            TestServices services = CreateServices(scenario.Application);
            SetSnapshot(
                services.ExcelInteropService,
                scenario.CaseWorkbook,
                "TASKPANE_BASE_SNAPSHOT_COUNT",
                "TASKPANE_BASE_SNAPSHOT_",
                embeddedSnapshot);
            services.ExcelInteropService.SetDocumentProperty(scenario.CaseWorkbook, "TASKPANE_BASE_MASTER_VERSION", "42");

            bool resolved = services.CaseCacheReader.TryEnsurePromotedCaseCacheThenResolve(scenario.CaseWorkbook, "1", out DocumentTemplateLookupResult lookupResult);

            Assert.True(resolved);
            Assert.NotNull(lookupResult);
            Assert.Equal("01", lookupResult.Key);
            Assert.Equal("委任状", lookupResult.DocumentName);
            Assert.Equal("01_委任状.docx", lookupResult.TemplateFileName);
            Assert.Equal(DocumentTemplateResolutionSource.SnapshotCache, lookupResult.ResolutionSource);
            Assert.Equal(embeddedSnapshot, scenario.LoadCaseCacheSnapshot());
            Assert.Equal("42", scenario.GetCaseProperty("TASKPANE_MASTER_VERSION"));
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

        [Fact]
        public void Resolve_WhenProcessHandlesMultipleSystemRoots_UsesMatchingMasterCatalogForEachRoot()
        {
            var application = new Excel.Application();
            using var rootAScenario = SnapshotBuilderScenario.Create(
                CreateRows("01_委任状A.docx", "委任状A"),
                masterVersion: 42,
                caseListRegistered: false,
                application: application,
                systemRoot: @"C:\SnapshotRegression\RootA",
                caseWorkbookPath: @"C:\SnapshotRegression\RootA\Cases\案件情報_A.xlsx",
                masterWorkbookPath: @"C:\SnapshotRegression\RootA\案件情報System_Kernel.xlsx");
            using var rootBScenario = SnapshotBuilderScenario.Create(
                CreateRows("01_委任状B.docx", "委任状B"),
                masterVersion: 42,
                caseListRegistered: false,
                application: application,
                systemRoot: @"C:\SnapshotRegression\RootB",
                caseWorkbookPath: @"C:\SnapshotRegression\RootB\Cases\案件情報_B.xlsx",
                masterWorkbookPath: @"C:\SnapshotRegression\RootB\案件情報System_Kernel.xlsx");
            EnsureMasterWorkbookFileExists(rootAScenario.MasterWorkbook.FullName);
            EnsureMasterWorkbookFileExists(rootBScenario.MasterWorkbook.FullName);

            TestServices services = CreateServices(application);

            DocumentTemplateSpec rootATemplate = services.Resolver.Resolve(rootAScenario.CaseWorkbook, "1");
            DocumentTemplateSpec rootBTemplate = services.Resolver.Resolve(rootBScenario.CaseWorkbook, "1");

            Assert.NotNull(rootATemplate);
            Assert.Equal("委任状A", rootATemplate.DocumentName);
            Assert.Equal("01_委任状A.docx", rootATemplate.TemplateFileName);
            Assert.Equal(DocumentTemplateResolutionSource.MasterCatalog, rootATemplate.ResolutionSource);
            Assert.NotNull(rootBTemplate);
            Assert.Equal("委任状B", rootBTemplate.DocumentName);
            Assert.Equal("01_委任状B.docx", rootBTemplate.TemplateFileName);
            Assert.Equal(DocumentTemplateResolutionSource.MasterCatalog, rootBTemplate.ResolutionSource);
        }

        [Fact]
        public void InvalidateCache_WhenWorkbookProvided_RefreshesOnlyMatchingMasterPath()
        {
            var application = new Excel.Application();
            using var rootAScenario = SnapshotBuilderScenario.Create(
                CreateRows("01_委任状A.docx", "委任状A"),
                masterVersion: 42,
                caseListRegistered: false,
                application: application,
                systemRoot: @"C:\SnapshotRegression\RootA",
                caseWorkbookPath: @"C:\SnapshotRegression\RootA\Cases\案件情報_A.xlsx",
                masterWorkbookPath: @"C:\SnapshotRegression\RootA\案件情報System_Kernel.xlsx");
            using var rootBScenario = SnapshotBuilderScenario.Create(
                CreateRows("01_委任状B.docx", "委任状B"),
                masterVersion: 42,
                caseListRegistered: false,
                application: application,
                systemRoot: @"C:\SnapshotRegression\RootB",
                caseWorkbookPath: @"C:\SnapshotRegression\RootB\Cases\案件情報_B.xlsx",
                masterWorkbookPath: @"C:\SnapshotRegression\RootB\案件情報System_Kernel.xlsx");
            EnsureMasterWorkbookFileExists(rootAScenario.MasterWorkbook.FullName);
            EnsureMasterWorkbookFileExists(rootBScenario.MasterWorkbook.FullName);

            TestServices services = CreateServices(application);

            DocumentTemplateSpec beforeInvalidateRootA = services.Resolver.Resolve(rootAScenario.CaseWorkbook, "1");
            DocumentTemplateSpec beforeInvalidateRootB = services.Resolver.Resolve(rootBScenario.CaseWorkbook, "1");

            UpdateMasterTemplateRow(rootAScenario.MasterWorkbook, "1", "01_委任状A_改訂.docx", "委任状A改訂");
            UpdateMasterTemplateRow(rootBScenario.MasterWorkbook, "1", "01_委任状B_改訂.docx", "委任状B改訂");

            services.MasterTemplateCatalogService.InvalidateCache(rootAScenario.CaseWorkbook);

            DocumentTemplateSpec afterInvalidateRootA = services.Resolver.Resolve(rootAScenario.CaseWorkbook, "1");
            DocumentTemplateSpec afterInvalidateRootB = services.Resolver.Resolve(rootBScenario.CaseWorkbook, "1");

            Assert.NotNull(beforeInvalidateRootA);
            Assert.Equal("委任状A", beforeInvalidateRootA.DocumentName);
            Assert.NotNull(beforeInvalidateRootB);
            Assert.Equal("委任状B", beforeInvalidateRootB.DocumentName);
            Assert.NotNull(afterInvalidateRootA);
            Assert.Equal("委任状A改訂", afterInvalidateRootA.DocumentName);
            Assert.Equal("01_委任状A_改訂.docx", afterInvalidateRootA.TemplateFileName);
            Assert.NotNull(afterInvalidateRootB);
            Assert.Equal("委任状B", afterInvalidateRootB.DocumentName);
            Assert.Equal("01_委任状B.docx", afterInvalidateRootB.TemplateFileName);
        }

        private static TestServices CreateServices(Excel.Application application, List<string> logs = null)
        {
            var logger = new Logger(message =>
            {
                if (logs != null)
                {
                    logs.Add(message ?? string.Empty);
                }
            });
            var pathCompatibilityService = new PathCompatibilityService();
            var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);
            var taskPaneSnapshotCacheService = new TaskPaneSnapshotCacheService(excelInteropService, logger);
            var masterTemplateSheetReader = new MasterTemplateSheetReaderAdapter();
            var masterWorkbookReadAccessService = new MasterWorkbookReadAccessService(application, excelInteropService, pathCompatibilityService);
            var masterTemplateCatalogService = new MasterTemplateCatalogService(excelInteropService, masterWorkbookReadAccessService, masterTemplateSheetReader, logger);
            var lookupService = new DocumentTemplateLookupService(taskPaneSnapshotCacheService, masterTemplateCatalogService);

            return new TestServices(
                excelInteropService,
                masterTemplateCatalogService,
                lookupService,
                new DocumentTemplateResolver(excelInteropService, pathCompatibilityService, lookupService, logger),
                new DocumentNamePromptService(excelInteropService, lookupService, logger));
        }

        private static bool ContainsOrdinal(string text, string value)
        {
            return (text ?? string.Empty).IndexOf(value ?? string.Empty, System.StringComparison.Ordinal) >= 0;
        }

        private static SnapshotBuilderScenario.InputRow[] CreateRows()
        {
            return CreateRows("01_委任状.docx", "委任状");
        }

        private static SnapshotBuilderScenario.InputRow[] CreateRows(string templateFileName, string caption)
        {
            return new[]
            {
                new SnapshotBuilderScenario.InputRow
                {
                    Key = "1",
                    TemplateFileName = templateFileName,
                    Caption = caption,
                    TabName = "申請手続",
                    FillColor = 1111,
                    TabBackColor = 2222
                }
            };
        }

        private static void UpdateMasterTemplateRow(Excel.Workbook masterWorkbook, string key, string templateFileName, string caption)
        {
            Excel.Worksheet worksheet = masterWorkbook?.Worksheets["雛形一覧"];
            if (worksheet == null)
            {
                return;
            }

            worksheet.Cells[3, "A"].Value2 = key;
            worksheet.Cells[3, "B"].Value2 = templateFileName;
            worksheet.Cells[3, "C"].Value2 = caption;
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

        private static void SetSnapshot(
            ExcelInteropService excelInteropService,
            Excel.Workbook workbook,
            string countPropertyName,
            string partPropertyPrefix,
            string snapshotText)
        {
            const int ChunkSize = 240;
            string safeSnapshot = snapshotText ?? string.Empty;
            int partCount = safeSnapshot.Length == 0 ? 0 : ((safeSnapshot.Length - 1) / ChunkSize) + 1;
            excelInteropService.SetDocumentProperty(workbook, countPropertyName, partCount.ToString(CultureInfo.InvariantCulture));

            for (int index = 1; index <= partCount; index++)
            {
                int startIndex = (index - 1) * ChunkSize;
                int length = System.Math.Min(ChunkSize, safeSnapshot.Length - startIndex);
                excelInteropService.SetDocumentProperty(
                    workbook,
                    partPropertyPrefix + index.ToString("00", CultureInfo.InvariantCulture),
                    safeSnapshot.Substring(startIndex, length));
            }
        }

        private sealed class TestServices
        {
            internal TestServices(
                ExcelInteropService excelInteropService,
                MasterTemplateCatalogService masterTemplateCatalogService,
                ICaseCacheDocumentTemplateReader caseCacheReader,
                DocumentTemplateResolver resolver,
                DocumentNamePromptService promptService)
            {
                ExcelInteropService = excelInteropService;
                MasterTemplateCatalogService = masterTemplateCatalogService;
                CaseCacheReader = caseCacheReader;
                Resolver = resolver;
                PromptService = promptService;
            }

            internal ExcelInteropService ExcelInteropService { get; }

            internal MasterTemplateCatalogService MasterTemplateCatalogService { get; }

            internal ICaseCacheDocumentTemplateReader CaseCacheReader { get; }

            internal DocumentTemplateResolver Resolver { get; }

            internal DocumentNamePromptService PromptService { get; }
        }
    }
}
