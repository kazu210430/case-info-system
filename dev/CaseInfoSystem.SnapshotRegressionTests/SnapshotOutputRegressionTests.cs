using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    public class SnapshotOutputRegressionTests
    {
        [Fact]
        public void BuildSnapshotText_MasterListRebuild_MatchesLegacySnapshotTextExactly()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult result = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            string expected = SnapshotLegacySerializer.Serialize(
                scenario.CaseWorkbook,
                42,
                scenario.Values,
                scenario.FillColors,
                scenario.TabBackColors);

            Assert.True(result.UpdatedCaseSnapshotCache);
            Assert.Equal(expected, result.SnapshotText);
            Assert.Equal(expected, scenario.LoadCaseCacheSnapshot());
            Assert.Equal("42", scenario.GetCaseProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void BuildSnapshotText_RebuildThenCaseCache_ProducesStableProjectionJson()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: true);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult first = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult second = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            string expected = SnapshotLegacySerializer.Serialize(
                scenario.CaseWorkbook,
                42,
                scenario.Values,
                scenario.FillColors,
                scenario.TabBackColors);

            Assert.True(first.UpdatedCaseSnapshotCache);
            Assert.False(second.UpdatedCaseSnapshotCache);
            Assert.Equal(TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.CaseCache, second.SnapshotSource);
            Assert.Equal(expected, first.SnapshotText);
            Assert.Equal(expected, second.SnapshotText);
            Assert.Equal(
                SnapshotProjection.FromSnapshotText(expected).ToJson(),
                SnapshotProjection.FromSnapshotText(second.SnapshotText).ToJson());
        }

        [Fact]
        public void BuildSnapshotText_WhenCaseMasterVersionMatchesLatest_UsesBaseSnapshotReadOnlyWithoutSavingCaseCache()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: true);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);
            string expected = SnapshotLegacySerializer.Serialize(
                scenario.CaseWorkbook,
                42,
                scenario.Values,
                scenario.FillColors,
                scenario.TabBackColors);
            scenario.SetCaseProperty("TASKPANE_MASTER_VERSION", "42");
            scenario.SetCaseProperty("TASKPANE_BASE_MASTER_VERSION", "42");
            scenario.SetCaseProperty("TASKPANE_BASE_SNAPSHOT_COUNT", "1");
            scenario.SetCaseProperty("TASKPANE_BASE_SNAPSHOT_01", expected);

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult result = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);

            Assert.False(result.UpdatedCaseSnapshotCache);
            Assert.Equal(TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.BaseCache, result.SnapshotSource);
            Assert.Equal(expected, result.SnapshotText);
            Assert.Equal(string.Empty, scenario.LoadCaseCacheSnapshot());
            Assert.Equal(string.Empty, scenario.GetCaseProperty("TASKPANE_SNAPSHOT_CACHE_COUNT"));
            Assert.Equal("42", scenario.GetCaseProperty("TASKPANE_MASTER_VERSION"));
        }

        [Fact]
        public void BuildSnapshotText_WhenCaseMasterVersionDiffers_RebuildsAndUpdatesCaseVersion()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);
            scenario.SetCaseProperty("TASKPANE_MASTER_VERSION", "41");

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult result = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);

            Assert.True(result.UpdatedCaseSnapshotCache);
            Assert.Equal(TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.MasterListRebuild, result.SnapshotSource);
            Assert.Equal("42", scenario.GetCaseProperty("TASKPANE_MASTER_VERSION"));
            Assert.Equal(result.SnapshotText, scenario.LoadCaseCacheSnapshot());
        }

        [Fact]
        public void BuildSnapshotText_PreservesCurrentRowHandling_ForEmptyBlankTabAndIncompleteRows()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult result = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);
            SnapshotProjection projection = SnapshotProjection.FromSnapshotText(result.SnapshotText);

            Assert.Equal(
                "{\"exportVersion\":\"2\",\"masterVersion\":\"42\",\"workbookName\":\"案件情報_山田.xlsx\",\"workbookPath\":\"C:\\\\SnapshotRegression\\\\Cases\\\\案件情報_山田.xlsx\",\"preferredPaneWidth\":420,\"specialButtons\":[\"btnCaseList|案件一覧登録（未了）|caselist||18|16|128|32|14803448\",\"btnAccounting|会計書類セット|accounting||18|64|128|32|14348250\"],\"tabs\":[\"1|申請手続|2222\",\"2|その他|0\",\"3|契約確認|7777\",\"4|全て|16777215\"],\"docs\":[\"btnDoc_01|01|委任状|doc|申請手続|1|1111|01_委任状.docx\",\"btnDoc_02|02|見積書|doc|その他|1|3333|02_見積書.dotx\",\"btnDoc_10|10|報告書|doc|申請手続|2|4444|10_報告書.dotm\",\"btnDoc_11|11|不完全行|doc|契約確認|1|6666|\",\"btnDoc_12|12|請求書|doc|契約確認|2|8888|12_請求書.docx\"]}",
                projection.ToJson());
        }

        [Fact]
        public void BuildSnapshotText_WhenMasterWorkbookAlreadyOpen_DoesNotChangeWindowVisibility()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);

            scenario.MasterWorkbook.Windows[1].Visible = true;

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult result = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);

            Assert.True(result.UpdatedCaseSnapshotCache);
            Assert.True(scenario.MasterWorkbook.Windows[1].Visible);
        }

        [Fact]
        public void BuildSnapshotText_WhenMasterWorkbookOpenedForRead_HidesOnlyOpenedWorkbook()
        {
            using var scenario = SnapshotBuilderScenario.Create(CreateRows(), masterVersion: 42, caseListRegistered: false);
            EnsureMasterWorkbookFileExists(scenario.MasterWorkbook.FullName);
            scenario.Application.Workbooks.Remove(scenario.MasterWorkbook);
            scenario.MasterWorkbook.Windows[1].Visible = true;
            scenario.Application.Workbooks.OpenBehavior = (_, __, ___) => scenario.MasterWorkbook;

            TaskPaneSnapshotBuilderService.TaskPaneBuildResult result = scenario.Builder.BuildSnapshotText(scenario.CaseWorkbook);

            Assert.True(result.UpdatedCaseSnapshotCache);
            Assert.False(scenario.MasterWorkbook.Windows[1].Visible);
            Assert.DoesNotContain(scenario.MasterWorkbook, scenario.Application.Workbooks);
        }

        private static IReadOnlyList<SnapshotBuilderScenario.InputRow> CreateRows()
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
                },
                new SnapshotBuilderScenario.InputRow(),
                new SnapshotBuilderScenario.InputRow
                {
                    Key = "2",
                    TemplateFileName = "02_見積書.dotx",
                    Caption = "見積書",
                    TabName = string.Empty,
                    FillColor = 3333,
                    TabBackColor = 0
                },
                new SnapshotBuilderScenario.InputRow
                {
                    Key = "10",
                    TemplateFileName = "10_報告書.dotm",
                    Caption = "報告書",
                    TabName = "申請手続",
                    FillColor = 4444,
                    TabBackColor = 5555
                },
                new SnapshotBuilderScenario.InputRow
                {
                    Key = "11",
                    TemplateFileName = string.Empty,
                    Caption = "不完全行",
                    TabName = "契約確認",
                    FillColor = 6666,
                    TabBackColor = 7777
                },
                new SnapshotBuilderScenario.InputRow
                {
                    Key = "12",
                    TemplateFileName = "12_請求書.docx",
                    Caption = "請求書",
                    TabName = "契約確認",
                    FillColor = 8888,
                    TabBackColor = 9999
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
    }
}
