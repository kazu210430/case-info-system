using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // Thin shell responsibilities:
    // 1. Route separated action kinds to dedicated handlers.
    // 2. Keep a single entry point into the frozen fallback path.
    internal sealed class TaskPaneActionDispatcher
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private const string DocumentActionKind = "doc";
        private const string AccountingActionKind = "accounting";

        private readonly ThisAddIn _addIn;
        private readonly ExcelInteropService _excelInteropService;
        private readonly TaskPaneCaseAccountingActionHandler _taskPaneCaseAccountingActionHandler;
        private readonly TaskPaneCaseDocumentActionHandler _taskPaneCaseDocumentActionHandler;
        private readonly TaskPaneCaseFallbackActionExecutor _taskPaneCaseFallbackActionExecutor;
        private readonly CaseTaskPaneViewStateBuilder _caseTaskPaneViewStateBuilder;
        private readonly UserErrorService _userErrorService;
        private readonly Logger _logger;
        private readonly TaskPaneCaseActionTargetResolver _caseActionTargetResolver;
        private readonly Action<TaskPaneHost> _invalidateHostRenderStateForForcedRefresh;
        private readonly Action<DocumentButtonsControl, Excel.Workbook> _renderCaseHostAfterAction;
        private readonly Func<TaskPaneHost, string, bool> _tryShowHost;

        internal TaskPaneActionDispatcher(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneCaseFallbackActionExecutor taskPaneCaseFallbackActionExecutor,
            TaskPaneCaseActionTargetResolver caseActionTargetResolver,
            TaskPaneCaseAccountingActionHandler taskPaneCaseAccountingActionHandler,
            TaskPaneCaseDocumentActionHandler taskPaneCaseDocumentActionHandler,
            Action<TaskPaneHost> invalidateHostRenderStateForForcedRefresh,
            Action<DocumentButtonsControl, Excel.Workbook> renderCaseHostAfterAction,
            Func<TaskPaneHost, string, bool> tryShowHost)
        {
            _addIn = addIn;
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _caseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            _userErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _taskPaneCaseFallbackActionExecutor = taskPaneCaseFallbackActionExecutor ?? throw new ArgumentNullException(nameof(taskPaneCaseFallbackActionExecutor));
            _caseActionTargetResolver = caseActionTargetResolver ?? throw new ArgumentNullException(nameof(caseActionTargetResolver));
            _taskPaneCaseAccountingActionHandler = taskPaneCaseAccountingActionHandler ?? throw new ArgumentNullException(nameof(taskPaneCaseAccountingActionHandler));
            _taskPaneCaseDocumentActionHandler = taskPaneCaseDocumentActionHandler ?? throw new ArgumentNullException(nameof(taskPaneCaseDocumentActionHandler));
            _invalidateHostRenderStateForForcedRefresh = invalidateHostRenderStateForForcedRefresh ?? throw new ArgumentNullException(nameof(invalidateHostRenderStateForForcedRefresh));
            _renderCaseHostAfterAction = renderCaseHostAfterAction ?? throw new ArgumentNullException(nameof(renderCaseHostAfterAction));
            _tryShowHost = tryShowHost ?? throw new ArgumentNullException(nameof(tryShowHost));
        }

        internal void HandleCaseControlActionInvoked(string windowKey, DocumentButtonsControl control, TaskPaneActionEventArgs e)
        {
            if (TryRouteSeparatedActionKind(windowKey, control, e))
            {
                return;
            }

            HandleFrozenFallbackActionEntry(windowKey, control, e);
        }

        private bool TryRouteSeparatedActionKind(string windowKey, DocumentButtonsControl control, TaskPaneActionEventArgs e)
        {
            string actionKind = e?.ActionKind;
            if (string.Equals(actionKind, DocumentActionKind, StringComparison.OrdinalIgnoreCase))
            {
                _taskPaneCaseDocumentActionHandler.HandleCaseControlActionInvoked(windowKey, control, e.Key);
                return true;
            }

            if (string.Equals(actionKind, AccountingActionKind, StringComparison.OrdinalIgnoreCase))
            {
                _taskPaneCaseAccountingActionHandler.HandleCaseControlActionInvoked(windowKey, control, e.Key);
                return true;
            }

            return false;
        }

        // Keep target resolution, exception handling, and refresh ordering behind one
        // frozen fallback entry until this path is intentionally split further.
        private void HandleFrozenFallbackActionEntry(string windowKey, DocumentButtonsControl control, TaskPaneActionEventArgs e)
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
                bool shouldContinue = _taskPaneCaseFallbackActionExecutor.TryExecute(workbook, e);
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

        // Preserve the existing post-action refresh/render/show ordering for the frozen fallback path.
        internal void HandlePostActionRefresh(TaskPaneHost host, Excel.Workbook workbook, DocumentButtonsControl control, string actionKind)
        {
            TaskPanePostActionRefreshDecision decision = TaskPanePostActionRefreshPolicy.Decide(actionKind);
            bool beforeSignaturePresent = host != null && !string.IsNullOrWhiteSpace(host.LastRenderSignature);
            bool addInPresent = _addIn != null;
            bool hostWindowPresent = host != null && host.Window != null;
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
                bool afterSignaturePresent = host != null && !string.IsNullOrWhiteSpace(host.LastRenderSignature);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneActionDispatcher action=post-action-metadata actionKind="
                    + (actionKind ?? string.Empty)
                    + ", postActionInvalidation=true"
                    + ", fallbackRewrite=false"
                    + ", fallbackReason=\"CASE pane refresh after case-list action was deferred so Kernel navigation can take the foreground.\""
                    + ", beforeSignaturePresent="
                    + beforeSignaturePresent.ToString()
                    + ", afterSignaturePresent="
                    + afterSignaturePresent.ToString()
                    + ", invalidateThenLocalRender=false"
                    + ", tryShowHostAfterRewrite=false"
                    + ", addInPresent="
                    + addInPresent.ToString()
                    + ", hostWindowPresent="
                    + hostWindowPresent.ToString());
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

            bool beforeSignaturePresent = !string.IsNullOrWhiteSpace(host.LastRenderSignature);
            bool addInPresent = _addIn != null;
            bool hostWindowPresent = host.Window != null;
            if (addInPresent && hostWindowPresent)
            {
                _addIn.RequestTaskPaneDisplayForTargetWindow(
                    TaskPaneDisplayRequest.ForPostActionRefresh(actionKind),
                    workbook,
                    host.Window);
                return;
            }

            string fallbackReason = !addInPresent
                ? "add-in was not available for post-action display request."
                : "host window was not available for post-action display request.";
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
            bool afterSignaturePresent = !string.IsNullOrWhiteSpace(host.LastRenderSignature);
            bool tryShowHostAfterRewrite = _tryShowHost(host, "RefreshCaseHostAfterAction");
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneActionDispatcher action=post-action-metadata actionKind="
                + (actionKind ?? string.Empty)
                + ", postActionInvalidation=true"
                + ", fallbackRewrite=true"
                + ", fallbackReason=\""
                + fallbackReason
                + "\""
                + ", beforeSignaturePresent="
                + beforeSignaturePresent.ToString()
                + ", afterSignaturePresent="
                + afterSignaturePresent.ToString()
                + ", invalidateThenLocalRender=true"
                + ", tryShowHostAfterRewrite="
                + tryShowHostAfterRewrite.ToString()
                + ", addInPresent="
                + addInPresent.ToString()
                + ", hostWindowPresent="
                + hostWindowPresent.ToString());
            if (!tryShowHostAfterRewrite)
            {
                _logger.Warn("CASE pane refresh after action skipped because host could not be shown. workbook=" + (host.WorkbookFullName ?? string.Empty));
                return;
            }

            _logger.Info("CASE pane refreshed after action. workbook=" + (host.WorkbookFullName ?? string.Empty));
        }
    }
}
