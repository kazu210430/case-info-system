using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneCaseAccountingActionHandler
    {
        private const string AccountingActionKind = "accounting";

        private readonly TaskPaneCaseActionTargetResolver _caseActionTargetResolver;
        private readonly TaskPaneCaseFallbackActionExecutor _taskPaneCaseFallbackActionExecutor;
        private readonly CaseTaskPaneViewStateBuilder _caseTaskPaneViewStateBuilder;
        private readonly UserErrorService _userErrorService;
        private readonly Logger _logger;
        private readonly Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> _handlePostActionRefresh;

        internal TaskPaneCaseAccountingActionHandler(
            TaskPaneCaseActionTargetResolver caseActionTargetResolver,
            TaskPaneCaseFallbackActionExecutor taskPaneCaseFallbackActionExecutor,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh)
        {
            _caseActionTargetResolver = caseActionTargetResolver ?? throw new ArgumentNullException(nameof(caseActionTargetResolver));
            _taskPaneCaseFallbackActionExecutor = taskPaneCaseFallbackActionExecutor ?? throw new ArgumentNullException(nameof(taskPaneCaseFallbackActionExecutor));
            _caseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            _userErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _handlePostActionRefresh = handlePostActionRefresh ?? throw new ArgumentNullException(nameof(handlePostActionRefresh));
        }

        internal void HandleCaseControlActionInvoked(string windowKey, DocumentButtonsControl control, string key)
        {
            if (control == null)
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because host identity was not available.");
                return;
            }

            if (!_caseActionTargetResolver.TryResolve(windowKey, out TaskPaneHost host, out Excel.Workbook workbook))
            {
                if (host != null && workbook == null)
                {
                    control.Render(_caseTaskPaneViewStateBuilder.BuildWorkbookNotFoundState());
                }

                return;
            }

            try
            {
                bool shouldContinue = _taskPaneCaseFallbackActionExecutor.TryExecute(workbook, AccountingActionKind, key);
                if (!shouldContinue)
                {
                    return;
                }

                _handlePostActionRefresh(host, workbook, control, AccountingActionKind);
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
