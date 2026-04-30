using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Core;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    public class TaskPaneSnapshotCacheStorageBehaviorTests
    {
        [Fact]
        public void PromoteBaseSnapshotToCaseCacheIfNeeded_WhenCaseCacheIsEmpty_PromotesEmbeddedSnapshotAndSetsCaseVersion()
        {
            using var context = TestContext.Create();
            string embeddedSnapshot = CreateCompatibleSnapshotText("01", "委任状", "01_委任状.docx");
            context.SetBaseSnapshot(embeddedSnapshot);
            context.SetProperty("TASKPANE_BASE_MASTER_VERSION", "7");

            bool promoted = context.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(context.Workbook);

            Assert.True(promoted);
            Assert.Equal(embeddedSnapshot, context.LoadCaseSnapshot());
            Assert.Equal("7", context.GetProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void PromoteBaseSnapshotToCaseCacheIfNeeded_WhenEmbeddedVersionIsNewer_PromotesOverExistingCaseCache()
        {
            using var context = TestContext.Create();
            string existingCaseSnapshot = CreateCompatibleSnapshotText("01", "CASE旧", "01_case-old.docx");
            string embeddedSnapshot = CreateCompatibleSnapshotText("02", "BASE新", "02_base-new.docx");
            context.SetCaseSnapshot(existingCaseSnapshot);
            context.SetBaseSnapshot(embeddedSnapshot);
            context.SetProperty("TASKPANE_MASTER_VERSION", "3");
            context.SetProperty("TASKPANE_BASE_MASTER_VERSION", "5");

            bool promoted = context.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(context.Workbook);

            Assert.True(promoted);
            Assert.Equal(embeddedSnapshot, context.LoadCaseSnapshot());
            Assert.Equal("5", context.GetProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void PromoteBaseSnapshotToCaseCacheIfNeeded_WhenCaseVersionIsMissingAndEmbeddedVersionExists_Promotes()
        {
            using var context = TestContext.Create();
            string existingCaseSnapshot = CreateCompatibleSnapshotText("01", "CASE既存", "01_case.docx");
            string embeddedSnapshot = CreateCompatibleSnapshotText("02", "BASE埋込", "02_base.docx");
            context.SetCaseSnapshot(existingCaseSnapshot);
            context.SetBaseSnapshot(embeddedSnapshot);
            context.SetProperty("TASKPANE_BASE_MASTER_VERSION", "8");

            bool promoted = context.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(context.Workbook);

            Assert.True(promoted);
            Assert.Equal(embeddedSnapshot, context.LoadCaseSnapshot());
            Assert.Equal("8", context.GetProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void PromoteBaseSnapshotToCaseCacheIfNeeded_WhenCaseCacheIsValidAndEmbeddedIsNotNewer_DoesNotPromote()
        {
            using var context = TestContext.Create();
            string existingCaseSnapshot = CreateCompatibleSnapshotText("01", "CASE維持", "01_case.docx");
            string embeddedSnapshot = CreateCompatibleSnapshotText("02", "BASE古い", "02_base.docx");
            context.SetCaseSnapshot(existingCaseSnapshot);
            context.SetBaseSnapshot(embeddedSnapshot);
            context.SetProperty("TASKPANE_MASTER_VERSION", "9");
            context.SetProperty("TASKPANE_BASE_MASTER_VERSION", "5");

            bool promoted = context.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(context.Workbook);

            Assert.False(promoted);
            Assert.Equal(existingCaseSnapshot, context.LoadCaseSnapshot());
            Assert.Equal("9", context.GetProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void PromoteBaseSnapshotToCaseCacheIfNeeded_WhenCaseSnapshotIsIncompatible_ClearsCaseCache()
        {
            using var context = TestContext.Create();
            context.SetCaseSnapshot("META\t1\tlegacy");

            bool promoted = context.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(context.Workbook);

            Assert.False(promoted);
            Assert.Equal("0", context.GetProperty("TASKPANE_SNAPSHOT_CACHE_COUNT"));
            Assert.Equal(string.Empty, context.GetProperty("TASKPANE_SNAPSHOT_CACHE_01"));
        }

        [Fact]
        public void PromoteBaseSnapshotToCaseCacheIfNeeded_WhenEmbeddedSnapshotIsIncompatible_ClearsBaseSnapshotWithoutTouchingValidCaseCache()
        {
            using var context = TestContext.Create();
            string existingCaseSnapshot = CreateCompatibleSnapshotText("01", "CASE維持", "01_case.docx");
            context.SetCaseSnapshot(existingCaseSnapshot);
            context.SetBaseSnapshot("META\t1\tlegacy");
            context.SetProperty("TASKPANE_MASTER_VERSION", "9");
            context.SetProperty("TASKPANE_BASE_MASTER_VERSION", "10");

            bool promoted = context.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(context.Workbook);

            Assert.False(promoted);
            Assert.Equal(existingCaseSnapshot, context.LoadCaseSnapshot());
            Assert.Equal("0", context.GetProperty("TASKPANE_BASE_SNAPSHOT_COUNT"));
            Assert.Equal(string.Empty, context.GetProperty("TASKPANE_BASE_SNAPSHOT_01"));
        }

        [Fact]
        public void InitialPromote_OverwritesExistingCaseCacheWhereLookupPromoteKeepsNewerCaseCache()
        {
            string existingCaseSnapshot = CreateCompatibleSnapshotText("01", "CASE維持", "01_case.docx");
            string embeddedSnapshot = CreateCompatibleSnapshotText("02", "BASE初期", "02_base.docx");

            using var lookupContext = TestContext.Create();
            lookupContext.SetCaseSnapshot(existingCaseSnapshot);
            lookupContext.SetBaseSnapshot(embeddedSnapshot);
            lookupContext.SetProperty("TASKPANE_MASTER_VERSION", "9");
            lookupContext.SetProperty("TASKPANE_BASE_MASTER_VERSION", "5");

            bool promotedByLookup = lookupContext.CacheService.PromoteBaseSnapshotToCaseCacheIfNeeded(lookupContext.Workbook);

            using var initializerContext = TestContext.Create();
            initializerContext.SetCaseSnapshot(existingCaseSnapshot);
            initializerContext.SetBaseSnapshot(embeddedSnapshot);
            initializerContext.SetProperty("TASKPANE_MASTER_VERSION", "9");
            initializerContext.SetProperty("TASKPANE_BASE_MASTER_VERSION", "5");

            initializerContext.CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache(initializerContext.Workbook);

            Assert.False(promotedByLookup);
            Assert.Equal(existingCaseSnapshot, lookupContext.LoadCaseSnapshot());
            Assert.Equal(embeddedSnapshot, initializerContext.LoadCaseSnapshot());
            Assert.Equal("5", initializerContext.GetProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void PromoteEmbeddedSnapshotToCaseCache_WhenEmbeddedSnapshotIsIncompatible_ClearsBaseAndCaseSnapshotChunks()
        {
            using var context = TestContext.Create();
            string existingCaseSnapshot = CreateCompatibleSnapshotText("01", "CASE既存", "01_case.docx");
            context.SetCaseSnapshot(existingCaseSnapshot);
            context.SetBaseSnapshot("META\t1\tlegacy");

            context.CaseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache(context.Workbook);

            Assert.Equal("0", context.GetProperty("TASKPANE_BASE_SNAPSHOT_COUNT"));
            Assert.Equal(string.Empty, context.GetProperty("TASKPANE_BASE_SNAPSHOT_01"));
            Assert.Equal("0", context.GetProperty("TASKPANE_SNAPSHOT_CACHE_COUNT"));
            Assert.Equal(string.Empty, context.GetProperty("TASKPANE_SNAPSHOT_CACHE_01"));
        }

        private static string CreateCompatibleSnapshotText(string key, string caption, string templateFileName)
        {
            using var scenario = SnapshotBuilderScenario.Create(
                new[]
                {
                    new SnapshotBuilderScenario.InputRow
                    {
                        Key = key,
                        TemplateFileName = templateFileName,
                        Caption = caption,
                        TabName = "申請手続",
                        FillColor = 1111,
                        TabBackColor = 2222
                    }
                },
                masterVersion: 42,
                caseListRegistered: false);

            return scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook).SnapshotText;
        }

        private sealed class TestContext : IDisposable
        {
            private const int ChunkSize = 240;
            private readonly ExcelInteropService _excelInteropService;

            private TestContext(
                Excel.Application application,
                ExcelInteropService excelInteropService,
                TaskPaneSnapshotCacheService cacheService,
                CaseTemplateSnapshotService caseTemplateSnapshotService,
                Excel.Workbook workbook)
            {
                Application = application;
                _excelInteropService = excelInteropService;
                CacheService = cacheService;
                CaseTemplateSnapshotService = caseTemplateSnapshotService;
                Workbook = workbook;
            }

            internal Excel.Application Application { get; }

            internal TaskPaneSnapshotCacheService CacheService { get; }

            internal CaseTemplateSnapshotService CaseTemplateSnapshotService { get; }

            internal Excel.Workbook Workbook { get; }

            internal static TestContext Create()
            {
                var application = new Excel.Application();
                var logger = new Logger(_ => { });
                var pathCompatibilityService = new PathCompatibilityService();
                var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);
                var workbook = new Excel.Workbook
                {
                    FullName = @"C:\SnapshotRegression\Tests\案件情報_山田.xlsx",
                    Name = "案件情報_山田.xlsx",
                    Path = @"C:\SnapshotRegression\Tests",
                    CustomDocumentProperties = new DocumentProperties()
                };

                application.Workbooks.Add(workbook);

                return new TestContext(
                    application,
                    excelInteropService,
                    new TaskPaneSnapshotCacheService(excelInteropService, logger),
                    new CaseTemplateSnapshotService(excelInteropService),
                    workbook);
            }

            internal void SetCaseSnapshot(string snapshotText)
            {
                SetSnapshot("TASKPANE_SNAPSHOT_CACHE_COUNT", "TASKPANE_SNAPSHOT_CACHE_", snapshotText);
            }

            internal void SetBaseSnapshot(string snapshotText)
            {
                SetSnapshot("TASKPANE_BASE_SNAPSHOT_COUNT", "TASKPANE_BASE_SNAPSHOT_", snapshotText);
            }

            internal string LoadCaseSnapshot()
            {
                return LoadSnapshot("TASKPANE_SNAPSHOT_CACHE_COUNT", "TASKPANE_SNAPSHOT_CACHE_");
            }

            internal string GetProperty(string propertyName)
            {
                return _excelInteropService.TryGetDocumentProperty(Workbook, propertyName);
            }

            internal void SetProperty(string propertyName, string value)
            {
                _excelInteropService.SetDocumentProperty(Workbook, propertyName, value);
            }

            private void SetSnapshot(string countPropName, string partPropPrefix, string snapshotText)
            {
                int previousCount = ReadPositiveIntProperty(countPropName);
                string safeSnapshot = snapshotText ?? string.Empty;
                if (safeSnapshot.Length == 0)
                {
                    _excelInteropService.SetDocumentProperty(Workbook, countPropName, "0");
                    for (int index = 1; index <= previousCount; index++)
                    {
                        _excelInteropService.SetDocumentProperty(Workbook, partPropPrefix + index.ToString("00", CultureInfo.InvariantCulture), string.Empty);
                    }

                    return;
                }

                int partCount = ((safeSnapshot.Length - 1) / ChunkSize) + 1;
                _excelInteropService.SetDocumentProperty(Workbook, countPropName, partCount.ToString(CultureInfo.InvariantCulture));
                for (int index = 1; index <= partCount; index++)
                {
                    int startIndex = (index - 1) * ChunkSize;
                    int length = Math.Min(ChunkSize, safeSnapshot.Length - startIndex);
                    _excelInteropService.SetDocumentProperty(
                        Workbook,
                        partPropPrefix + index.ToString("00", CultureInfo.InvariantCulture),
                        safeSnapshot.Substring(startIndex, length));
                }

                for (int index = partCount + 1; index <= previousCount; index++)
                {
                    _excelInteropService.SetDocumentProperty(Workbook, partPropPrefix + index.ToString("00", CultureInfo.InvariantCulture), string.Empty);
                }
            }

            private string LoadSnapshot(string countPropName, string partPropPrefix)
            {
                int partCount = ReadPositiveIntProperty(countPropName);
                if (partCount <= 0)
                {
                    return string.Empty;
                }

                var builder = new System.Text.StringBuilder();
                for (int index = 1; index <= partCount; index++)
                {
                    builder.Append(_excelInteropService.TryGetDocumentProperty(Workbook, partPropPrefix + index.ToString("00", CultureInfo.InvariantCulture)));
                }

                return builder.ToString();
            }

            private int ReadPositiveIntProperty(string propertyName)
            {
                string text = _excelInteropService.TryGetDocumentProperty(Workbook, propertyName);
                return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value > 0
                    ? value
                    : 0;
            }

            public void Dispose()
            {
                foreach (Excel.Workbook workbook in Application.Workbooks)
                {
                    workbook.Close(false, null, null);
                    break;
                }
            }
        }
    }
}
