using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneActionDispatcher
    {
        private readonly ThisAddIn _addIn;
        private readonly ExcelInteropService _excelInteropService;
        private readonly TaskPaneBusinessActionLauncher _taskPaneBusinessActionLauncher;
        private readonly TaskPaneCaseDocActionHandler _taskPaneCaseDocActionHandler;
        private readonly CaseTaskPaneViewStateBuilder _caseTaskPaneViewStateBuilder;
        private readonly UserErrorService _userErrorService;
        private readonly Logger _logger;
        private readonly Func<string, TaskPaneHost> _resolveHost;
        private readonly Action<TaskPaneHost> _invalidateHostRenderStateForForcedRefresh;
        private readonly Action<DocumentButtonsControl, Excel.Workbook> _renderCaseHostAfterAction;
        private readonly Func<TaskPaneHost, string, bool> _tryShowHost;

        internal TaskPaneActionDispatcher(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            Func<string, TaskPaneHost> resolveHost,
            Action<TaskPaneHost> invalidateHostRenderStateForForcedRefresh,
            Action<DocumentButtonsControl, Excel.Workbook> renderCaseHostAfterAction,
            Func<TaskPaneHost, string, bool> tryShowHost)
        {
            _addIn = addIn;
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _taskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher ?? throw new ArgumentNullException(nameof(taskPaneBusinessActionLauncher));
            _caseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            _userErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _resolveHost = resolveHost ?? throw new ArgumentNullException(nameof(resolveHost));
            _invalidateHostRenderStateForForcedRefresh = invalidateHostRenderStateForForcedRefresh ?? throw new ArgumentNullException(nameof(invalidateHostRenderStateForForcedRefresh));
            _renderCaseHostAfterAction = renderCaseHostAfterAction ?? throw new ArgumentNullException(nameof(renderCaseHostAfterAction));
            _tryShowHost = tryShowHost ?? throw new ArgumentNullException(nameof(tryShowHost));
            _taskPaneCaseDocActionHandler = new TaskPaneCaseDocActionHandler(
                _excelInteropService,
                _taskPaneBusinessActionLauncher,
                _caseTaskPaneViewStateBuilder,
                _userErrorService,
                _logger,
                _resolveHost,
                HandlePostActionRefresh);
        }

        internal void HandleCaseControlActionInvoked(string windowKey, DocumentButtonsControl control, TaskPaneActionEventArgs e)
        {
            if (string.Equals(e?.ActionKind, "doc", StringComparison.OrdinalIgnoreCase))
            {
                _taskPaneCaseDocActionHandler.HandleCaseControlActionInvoked(windowKey, control, e.Key);
                return;
            }

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
                bool shouldContinue = _taskPaneBusinessActionLauncher.TryExecute(workbook, e.ActionKind, e.Key);
                if (!shouldContinue)
                {
                    return;
                }

                HandlePostActionRefresh(host, workbook, control, e.ActionKind);
            }
            catch (Exception ex)
            {
                _logger.Error("CaseControl_ActionInvoked failed.", ex);
                control.Render(_caseTaskPaneViewStateBuilder.BuildActionFailedState());
                _userErrorService.ShowUserError("CaseControl_ActionInvoked", ex);
            }
        }

        private void HandlePostActionRefresh(TaskPaneHost host, Excel.Workbook workbook, DocumentButtonsControl control, string actionKind)
        {
            TaskPanePostActionRefreshDecision decision = TaskPanePostActionRefreshPolicy.Decide(actionKind);
            if (decision == TaskPanePostActionRefreshDecision.SkipForForegroundPreservation)
            {
                string reason = string.Equals(actionKind, "accounting", StringComparison.OrdinalIgnoreCase)
                    ? "accounting set should keep the generated workbook in the foreground."
                    : "document create should keep Word in the foreground.";
                _logger.Info("CASE pane refresh after action skipped because " + reason);
                return;
            }

            if (decision == TaskPanePostActionRefreshDecision.DeferAndInvalidateSignature)
            {
                _invalidateHostRenderStateForForcedRefresh(host);
                _logger.Info("CASE pane refresh after case-list action was deferred so Kernel navigation can take the foreground.");
                return;
            }

            RefreshCaseHostAfterAction(host, workbook, control, actionKind);
        }

        private void RefreshCaseHostAfterAction(TaskPaneHost host, Excel.Workbook workbook, DocumentButtonsControl control, string actionKind)
        {
            if (host == null || workbook == null || control == null)
            {
                return;
            }

            if (_addIn != null && host.Window != null)
            {
                _addIn.RequestTaskPaneDisplayForTargetWindow(
                    TaskPaneDisplayRequest.ForPostActionRefresh(actionKind),
                    workbook,
                    host.Window);
                return;
            }

            _invalidateHostRenderStateForForcedRefresh(host);
            _renderCaseHostAfterAction(control, workbook);
            host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(
                _excelInteropService,
                new WorkbookContext(
                    workbook,
                    host.Window,
                    WorkbookRole.Case,
                    _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT"),
                    _excelInteropService.GetWorkbookFullName(workbook),
                    _excelInteropService.GetActiveSheetCodeName(workbook)));
            if (!_tryShowHost(host, "RefreshCaseHostAfterAction"))
            {
                _logger.Warn("CASE pane refresh after action skipped because host could not be shown. workbook=" + (host.WorkbookFullName ?? string.Empty));
                return;
            }

            _logger.Info("CASE pane refreshed after action. workbook=" + (host.WorkbookFullName ?? string.Empty));
        }
    }
}
