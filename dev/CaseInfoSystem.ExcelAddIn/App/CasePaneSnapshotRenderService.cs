using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class CasePaneSnapshotRenderService
    {
        internal sealed class CasePaneSnapshotRenderResult
        {
            internal CasePaneSnapshotRenderResult(TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult, TaskPaneSnapshot snapshot)
            {
                BuildResult = buildResult;
                Snapshot = snapshot;
            }

            internal TaskPaneSnapshotBuilderService.TaskPaneBuildResult BuildResult { get; }

            internal TaskPaneSnapshot Snapshot { get; }
        }

        private readonly ICaseTaskPaneSnapshotReader _caseTaskPaneSnapshotReader;
        private readonly CaseTaskPaneViewStateBuilder _caseTaskPaneViewStateBuilder;

        internal CasePaneSnapshotRenderService(
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder)
        {
            _caseTaskPaneSnapshotReader = caseTaskPaneSnapshotReader ?? throw new ArgumentNullException(nameof(caseTaskPaneSnapshotReader));
            _caseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
        }

        internal CasePaneSnapshotRenderResult Render(DocumentButtonsControl control, Excel.Workbook workbook)
        {
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult = _caseTaskPaneSnapshotReader.BuildSnapshotText(workbook);
            TaskPaneSnapshot snapshot = TaskPaneSnapshotParser.Parse(buildResult.SnapshotText);
            Render(control, snapshot);
            return new CasePaneSnapshotRenderResult(buildResult, snapshot);
        }

        internal void RenderAfterAction(DocumentButtonsControl control, Excel.Workbook workbook)
        {
            string snapshotText = _caseTaskPaneSnapshotReader.BuildSnapshotText(workbook).SnapshotText;
            TaskPaneSnapshot snapshot = TaskPaneSnapshotParser.Parse(snapshotText);
            Render(control, snapshot);
        }

        private void Render(DocumentButtonsControl control, TaskPaneSnapshot snapshot)
        {
            CaseTaskPaneViewState viewState = _caseTaskPaneViewStateBuilder.Build(snapshot, control.SelectedTabName);
            control.Render(viewState);
        }
    }
}
