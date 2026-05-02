using System;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // Executes the non-routed case-action path that still flows through TaskPaneBusinessActionLauncher.
    internal sealed class TaskPaneCaseFallbackActionExecutor
    {
        private readonly TaskPaneBusinessActionLauncher _taskPaneBusinessActionLauncher;

        internal TaskPaneCaseFallbackActionExecutor(TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher)
        {
            _taskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher ?? throw new ArgumentNullException(nameof(taskPaneBusinessActionLauncher));
        }

        internal bool TryExecute(Excel.Workbook workbook, TaskPaneActionEventArgs e)
        {
            return _taskPaneBusinessActionLauncher.TryExecute(workbook, e.ActionKind, e.Key);
        }
    }
}
