using System;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneCaseResidualActionExecutor
    {
        private readonly TaskPaneBusinessActionLauncher _taskPaneBusinessActionLauncher;

        internal TaskPaneCaseResidualActionExecutor(TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher)
        {
            _taskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher ?? throw new ArgumentNullException(nameof(taskPaneBusinessActionLauncher));
        }

        internal bool Execute(Excel.Workbook workbook, TaskPaneActionEventArgs e)
        {
            return _taskPaneBusinessActionLauncher.TryExecute(workbook, e.ActionKind, e.Key);
        }
    }
}
