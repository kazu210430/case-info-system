using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneCaseDocActionHandler
    {
        private const string DocumentActionKind = "doc";

        private readonly ExcelInteropService _excelInteropService;
        private readonly TaskPaneBusinessActionLauncher _taskPaneBusinessActionLauncher;
        private readonly CaseTaskPaneViewStateBuilder _caseTaskPaneViewStateBuilder;
        private readonly UserErrorService _userErrorService;
        private readonly Logger _logger;
        private readonly Func<string, TaskPaneHost> _resolveHost;
        private readonly Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> _handlePostActionRefresh;

        internal TaskPaneCaseDocActionHandler(
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            Func<string, TaskPaneHost> resolveHost,
            Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _taskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher ?? throw new ArgumentNullException(nameof(taskPaneBusinessActionLauncher));
            _caseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            _userErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _resolveHost = resolveHost ?? throw new ArgumentNullException(nameof(resolveHost));
            _handlePostActionRefresh = handlePostActionRefresh ?? throw new ArgumentNullException(nameof(handlePostActionRefresh));
        }

        internal void HandleCaseControlActionInvoked(string windowKey, DocumentButtonsControl control, string key)
        {
            if (string.IsNullOrWhiteSpace(windowKey) || control == null)
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because host identity was not available.");
                return;
            }

            TaskPaneHost host = _resolveHost(windowKey);
            if (host == null)
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because host was not found. windowKey=" + windowKey);
                return;
            }

            Excel.Workbook workbook = _excelInteropService.FindOpenWorkbook(host.WorkbookFullName);
            if (workbook == null)
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because workbook was not found. windowKey=" + windowKey);
                control.Render(_caseTaskPaneViewStateBuilder.BuildWorkbookNotFoundState());
                return;
            }

            try
            {
                bool shouldContinue = _taskPaneBusinessActionLauncher.TryExecute(workbook, DocumentActionKind, key);
                if (!shouldContinue)
                {
                    return;
                }

                _handlePostActionRefresh(host, workbook, control, DocumentActionKind);
            }
            catch (Exception ex)
            {
                _logger.Error("CaseControl_ActionInvoked failed.", ex);
                control.Render(_caseTaskPaneViewStateBuilder.BuildActionFailedState());
                _userErrorService.ShowUserError("CaseControl_ActionInvoked", ex);
            }
        }
    }
}
