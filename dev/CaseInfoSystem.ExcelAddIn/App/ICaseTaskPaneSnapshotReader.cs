using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// CASE TaskPane 表示用 snapshot を取得する consumer 側の参照口。
    /// </summary>
    internal interface ICaseTaskPaneSnapshotReader
    {
        TaskPaneSnapshotBuilderService.TaskPaneBuildResult BuildSnapshotText(Excel.Workbook workbook);
    }
}
