using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Core;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    public class TaskPaneSnapshotChunkHelperTests
    {
        private const string CountPropName = "TASKPANE_SNAPSHOT_CACHE_COUNT";
        private const string PartPropPrefix = "TASKPANE_SNAPSHOT_CACHE_";
        private const string MasterVersionPropName = "TASKPANE_MASTER_VERSION";
        private const string BaseMasterVersionPropName = "TASKPANE_BASE_MASTER_VERSION";

        [Fact]
        public void LoadSnapshot_WhenCountIsZero_ReturnsEmpty()
        {
            using var context = TestContext.Create();
            context.SetProperty(CountPropName, "0");
            context.SetProperty(PartPropPrefix + "01", "stale");

            string snapshot = TaskPaneSnapshotChunkReadHelper.LoadSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix);

            Assert.Equal(string.Empty, snapshot);
        }

        [Fact]
        public void LoadSnapshot_WhenCountIsPositive_ConcatenatesChunksInOrder()
        {
            using var context = TestContext.Create();
            context.SetProperty(CountPropName, "3");
            context.SetProperty(PartPropPrefix + "01", "first-");
            context.SetProperty(PartPropPrefix + "02", "second-");
            context.SetProperty(PartPropPrefix + "03", "third");

            string snapshot = TaskPaneSnapshotChunkReadHelper.LoadSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix);

            Assert.Equal("first-second-third", snapshot);
        }

        [Fact]
        public void LoadSnapshot_WhenChunkIsMissing_TreatsMissingChunkAsEmptyString()
        {
            using var context = TestContext.Create();
            context.SetProperty(CountPropName, "3");
            context.SetProperty(PartPropPrefix + "01", "alpha");
            context.SetProperty(PartPropPrefix + "03", "omega");

            string snapshot = TaskPaneSnapshotChunkReadHelper.LoadSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix);

            Assert.Equal("alphaomega", snapshot);
        }

        [Fact]
        public void LoadSnapshot_WhenSnapshotTextLooksLegacy_ReturnsRawTextWithoutCompatibilityCheck()
        {
            using var context = TestContext.Create();
            context.SetProperty(CountPropName, "1");
            context.SetProperty(PartPropPrefix + "01", "META\t1\tlegacy");

            string snapshot = TaskPaneSnapshotChunkReadHelper.LoadSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix);

            Assert.Equal("META\t1\tlegacy", snapshot);
        }

        [Fact]
        public void SaveSnapshot_WhenSnapshotExceedsDefaultChunkSize_SplitsIntoChunksAndUpdatesCount()
        {
            using var context = TestContext.Create();
            string firstChunk = new string('A', 240);
            string secondChunk = "tail!";
            string snapshotText = firstChunk + secondChunk;

            TaskPaneSnapshotChunkStorageHelper.SaveSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix,
                snapshotText);

            Assert.Equal("2", context.GetProperty(CountPropName));
            Assert.Equal(firstChunk, context.GetProperty(PartPropPrefix + "01"));
            Assert.Equal(secondChunk, context.GetProperty(PartPropPrefix + "02"));
        }

        [Fact]
        public void SaveSnapshot_WhenSnapshotIsEmpty_SetsCountToZeroWithoutClearingExistingChunks()
        {
            using var context = TestContext.Create();
            context.SetProperty(CountPropName, "2");
            context.SetProperty(PartPropPrefix + "01", "keep-1");
            context.SetProperty(PartPropPrefix + "02", "keep-2");

            TaskPaneSnapshotChunkStorageHelper.SaveSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix,
                string.Empty);

            Assert.Equal("0", context.GetProperty(CountPropName));
            Assert.Equal("keep-1", context.GetProperty(PartPropPrefix + "01"));
            Assert.Equal("keep-2", context.GetProperty(PartPropPrefix + "02"));
            Assert.NotNull(context.GetDocumentProperty(PartPropPrefix + "01"));
            Assert.NotNull(context.GetDocumentProperty(PartPropPrefix + "02"));
        }

        [Fact]
        public void SaveSnapshot_WhenOverwritingWithSmallerSnapshot_EmptiesTrailingChunksWithoutDeletingProperties()
        {
            using var context = TestContext.Create();
            string firstChunk = new string('X', 240);
            string initialSnapshot = firstChunk + "second";
            string replacementSnapshot = "replacement";
            TaskPaneSnapshotChunkStorageHelper.SaveSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix,
                initialSnapshot);

            TaskPaneSnapshotChunkStorageHelper.SaveSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix,
                replacementSnapshot);

            Assert.Equal("1", context.GetProperty(CountPropName));
            Assert.Equal(replacementSnapshot, context.GetProperty(PartPropPrefix + "01"));
            Assert.Equal(string.Empty, context.GetProperty(PartPropPrefix + "02"));
            Assert.NotNull(context.GetDocumentProperty(PartPropPrefix + "02"));
        }

        [Fact]
        public void SaveSnapshot_WhenSnapshotTextLooksLegacy_DoesNotApplyCompatibilityOrVersionPolicy()
        {
            using var context = TestContext.Create();
            context.SetProperty(MasterVersionPropName, "9");
            context.SetProperty(BaseMasterVersionPropName, "11");

            TaskPaneSnapshotChunkStorageHelper.SaveSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix,
                "META\t1\tlegacy");

            Assert.Equal("1", context.GetProperty(CountPropName));
            Assert.Equal("META\t1\tlegacy", context.GetProperty(PartPropPrefix + "01"));
            Assert.Equal("9", context.GetProperty(MasterVersionPropName));
            Assert.Equal("11", context.GetProperty(BaseMasterVersionPropName));
        }

        [Fact]
        public void ClearSnapshot_WhenChunksExist_SetsCountToZeroClearsChunkValuesAndLeavesOtherPropertiesUntouched()
        {
            using var context = TestContext.Create();
            context.SetProperty(CountPropName, "2");
            context.SetProperty(PartPropPrefix + "01", "chunk-1");
            context.SetProperty(PartPropPrefix + "02", "chunk-2");
            context.SetProperty(MasterVersionPropName, "7");
            context.SetProperty(BaseMasterVersionPropName, "12");

            TaskPaneSnapshotChunkStorageHelper.ClearSnapshot(
                context.ExcelInteropService,
                context.Workbook,
                CountPropName,
                PartPropPrefix);

            Assert.Equal("0", context.GetProperty(CountPropName));
            Assert.Equal(string.Empty, context.GetProperty(PartPropPrefix + "01"));
            Assert.Equal(string.Empty, context.GetProperty(PartPropPrefix + "02"));
            Assert.NotNull(context.GetDocumentProperty(PartPropPrefix + "01"));
            Assert.NotNull(context.GetDocumentProperty(PartPropPrefix + "02"));
            Assert.Equal("7", context.GetProperty(MasterVersionPropName));
            Assert.Equal("12", context.GetProperty(BaseMasterVersionPropName));
        }

        private sealed class TestContext : IDisposable
        {
            private TestContext(Excel.Application application, ExcelInteropService excelInteropService, Excel.Workbook workbook)
            {
                Application = application;
                ExcelInteropService = excelInteropService;
                Workbook = workbook;
            }

            internal Excel.Application Application { get; }

            internal ExcelInteropService ExcelInteropService { get; }

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

                return new TestContext(application, excelInteropService, workbook);
            }

            internal string GetProperty(string propertyName)
            {
                return ExcelInteropService.TryGetDocumentProperty(Workbook, propertyName);
            }

            internal DocumentProperty GetDocumentProperty(string propertyName)
            {
                return (Workbook.CustomDocumentProperties as DocumentProperties)?[propertyName];
            }

            internal void SetProperty(string propertyName, string value)
            {
                ExcelInteropService.SetDocumentProperty(Workbook, propertyName, value);
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
