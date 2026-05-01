using System.Reflection;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class CasePaneSnapshotRenderServiceTests
    {
        [Fact]
        public void Render_ReturnsBuildResultAndParsedSnapshot()
        {
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(CreateSnapshotText(), updatedCaseSnapshotCache: true);
            var snapshotReader = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook => buildResult
            };
            var service = new CasePaneSnapshotRenderService(snapshotReader, new CaseTaskPaneViewStateBuilder());
            var control = new DocumentButtonsControl();

            CasePaneSnapshotRenderService.CasePaneSnapshotRenderResult result = service.Render(control, new Excel.Workbook());

            Assert.Same(buildResult, result.BuildResult);
            Assert.False(result.Snapshot.HasError);
            Assert.Equal(2, result.Snapshot.Tabs.Count);
            Assert.Equal(2, result.Snapshot.DocButtons.Count);
            Assert.Equal("全て", control.SelectedTabName);
        }

        [Fact]
        public void RenderAfterAction_PreservesSelectedTabName()
        {
            var snapshotReader = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook => new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(CreateSnapshotText(), updatedCaseSnapshotCache: false)
            };
            var service = new CasePaneSnapshotRenderService(snapshotReader, new CaseTaskPaneViewStateBuilder());
            var control = new DocumentButtonsControl();

            service.Render(control, new Excel.Workbook());
            SelectTab(control, "個別");

            service.RenderAfterAction(control, new Excel.Workbook());

            Assert.Equal("個別", control.SelectedTabName);
        }

        private static void SelectTab(DocumentButtonsControl control, string tabName)
        {
            FieldInfo innerControlField = typeof(DocumentButtonsControl).GetField("_innerControl", BindingFlags.Instance | BindingFlags.NonPublic);
            var innerControl = (DocTaskPaneControl)innerControlField.GetValue(control);
            FieldInfo currentViewStateField = typeof(DocTaskPaneControl).GetField("_currentViewState", BindingFlags.Instance | BindingFlags.NonPublic);
            var currentViewState = (CaseTaskPaneViewState)currentViewStateField.GetValue(innerControl);
            innerControl.Render(currentViewState.WithSelectedTab(tabName));
        }

        private static string CreateSnapshotText()
        {
            return string.Join(
                "\n",
                "META\tCASE\tcase.xlsx\tC:\\cases\\case.xlsx\t320",
                "TAB\t1\t全て\t16777215",
                "TAB\t2\t個別\t16777215",
                "DOC\tdoc01\t01\t文書A\tdoc\t全て\t1\t16777215\tA.docx",
                "DOC\tdoc02\t02\t文書B\tdoc\t個別\t2\t16777215\tB.docx");
        }
    }
}
