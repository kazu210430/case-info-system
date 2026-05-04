using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneHostFlowService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly ExcelInteropService _excelInteropService;
        private readonly TaskPaneDisplayCoordinator _taskPaneDisplayCoordinator;
        private readonly TaskPaneHostLifecycleService _taskPaneHostLifecycleService;
        private readonly Logger _logger;
        private readonly Func<WorkbookContext, string> _formatContextDescriptor;
        private readonly Func<TaskPaneHost, string> _formatHostDescriptor;
        private readonly Func<Excel.Window, string> _safeGetWindowKey;
        private readonly Action<TaskPaneHost, WorkbookContext, string> _renderHost;
        private int _kernelFlickerTraceRefreshPaneSequence;

        internal TaskPaneHostFlowService(
            ExcelInteropService excelInteropService,
            TaskPaneDisplayCoordinator taskPaneDisplayCoordinator,
            TaskPaneHostLifecycleService taskPaneHostLifecycleService,
            Logger logger,
            Func<WorkbookContext, string> formatContextDescriptor,
            Func<TaskPaneHost, string> formatHostDescriptor,
            Func<Excel.Window, string> safeGetWindowKey,
            Action<TaskPaneHost, WorkbookContext, string> renderHost)
        {
            _excelInteropService = excelInteropService;
            _taskPaneDisplayCoordinator = taskPaneDisplayCoordinator ?? throw new ArgumentNullException(nameof(taskPaneDisplayCoordinator));
            _taskPaneHostLifecycleService = taskPaneHostLifecycleService ?? throw new ArgumentNullException(nameof(taskPaneHostLifecycleService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _formatContextDescriptor = formatContextDescriptor ?? throw new ArgumentNullException(nameof(formatContextDescriptor));
            _formatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
            _safeGetWindowKey = safeGetWindowKey ?? throw new ArgumentNullException(nameof(safeGetWindowKey));
            _renderHost = renderHost ?? throw new ArgumentNullException(nameof(renderHost));
        }

        internal bool RefreshPane(WorkbookContext context, string reason)
        {
            int refreshPaneCallId = ++_kernelFlickerTraceRefreshPaneSequence;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=refresh-pane-start refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", reason="
                + (reason ?? string.Empty)
                + ", context="
                + _formatContextDescriptor(context));
            if (!TryAcceptRefreshPaneRequest(context, reason, refreshPaneCallId, out WorkbookRole role, out string windowKey))
            {
                return false;
            }

            _taskPaneHostLifecycleService.RemoveStaleKernelHostsForRefresh(context, windowKey);
            TaskPaneHost host = _taskPaneHostLifecycleService.GetOrReplaceHost(windowKey, context.Window, role);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=host-selected refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", host="
                + _formatHostDescriptor(host));
            if (TryReuseCaseHostForRefresh(host, context, reason, windowKey, refreshPaneCallId))
            {
                return true;
            }

            return RenderAndShowHostForRefresh(host, context, reason, windowKey, refreshPaneCallId, role);
        }

        private bool TryAcceptRefreshPaneRequest(WorkbookContext context, string reason, int refreshPaneCallId, out WorkbookRole role, out string windowKey)
        {
            role = context == null ? WorkbookRole.Unknown : context.Role;
            windowKey = string.Empty;
            if (TaskPaneRefreshPreconditionPolicy.ShouldHideAllAndSkip(role, windowKey: null))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=hide-all refreshPaneCallId="
                    + refreshPaneCallId.ToString()
                    + ", reason=PreconditionPolicyRole"
                    + ", role="
                    + role.ToString());
                _taskPaneDisplayCoordinator.HideAll();
                return false;
            }

            windowKey = _safeGetWindowKey(context == null ? null : context.Window);
            if (TaskPaneRefreshPreconditionPolicy.ShouldHideAllAndSkip(role, windowKey))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=hide-all refreshPaneCallId="
                    + refreshPaneCallId.ToString()
                    + ", reason=PreconditionPolicyWindowKey"
                    + ", role="
                    + role.ToString()
                    + ", windowKey="
                    + windowKey);
                _taskPaneDisplayCoordinator.HideAll();
                _logger.Warn("RefreshPane skipped because windowKey was empty. reason=" + (reason ?? string.Empty));
                return false;
            }

            return true;
        }

        private bool TryReuseCaseHostForRefresh(TaskPaneHost host, WorkbookContext context, string reason, string windowKey, int refreshPaneCallId)
        {
            if (!ShouldReuseCaseHostWithoutRender(host, context, reason))
            {
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=reuse-case-host refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", host="
                + _formatHostDescriptor(host)
                + ", reason="
                + (reason ?? string.Empty));
            _taskPaneDisplayCoordinator.PrepareHostsBeforeShow(host);
            if (!_taskPaneDisplayCoordinator.TryShowHost(host, "RefreshPane.ReuseCaseHost"))
            {
                _logger.Warn("RefreshPane skipped because reused CASE host could not be shown. reason=" + (reason ?? string.Empty) + ", windowKey=" + windowKey);
                return false;
            }

            _logger.Info("TaskPane reused. reason=" + (reason ?? string.Empty) + ", role=" + context.Role + ", windowKey=" + windowKey);
            return true;
        }

        private bool RenderAndShowHostForRefresh(TaskPaneHost host, WorkbookContext context, string reason, string windowKey, int refreshPaneCallId, WorkbookRole role)
        {
            TaskPaneRenderStateEvaluation renderState = TaskPaneRenderStateEvaluator.EvaluateRenderState(
                _excelInteropService,
                host,
                context);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=render-evaluate refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", host="
                + _formatHostDescriptor(host)
                + ", renderRequired="
                + renderState.IsRenderRequired.ToString());
            if (renderState.IsRenderRequired)
            {
                _renderHost(host, context, reason);
                host.LastRenderSignature = renderState.RenderSignature;
            }
            else
            {
                _logger.Debug(nameof(TaskPaneManager), "RefreshPane render skipped because the host state did not change. windowKey=" + windowKey + ", role=" + role);
            }

            _taskPaneDisplayCoordinator.PrepareHostsBeforeShow(host);
            if (!_taskPaneDisplayCoordinator.TryShowHost(host, "RefreshPane"))
            {
                _logger.Warn("RefreshPane skipped because host could not be shown. reason=" + (reason ?? string.Empty) + ", windowKey=" + windowKey);
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=refresh-pane-end refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", host="
                + _formatHostDescriptor(host)
                + ", result=Shown");
            _logger.Info("TaskPane refreshed. reason=" + (reason ?? string.Empty) + ", role=" + role + ", windowKey=" + windowKey);
            return true;
        }

        private static bool ShouldReuseCaseHostWithoutRender(TaskPaneHost host, WorkbookContext context, string reason)
        {
            if (host == null || context == null)
            {
                return false;
            }

            return TaskPaneHostReusePolicy.ShouldReuseCaseHostWithoutRender(
                context.Role,
                host.Control is DocumentButtonsControl,
                !string.IsNullOrWhiteSpace(host.LastRenderSignature),
                string.Equals(host.WorkbookFullName, context.WorkbookFullName, StringComparison.OrdinalIgnoreCase),
                reason);
        }
    }
}
