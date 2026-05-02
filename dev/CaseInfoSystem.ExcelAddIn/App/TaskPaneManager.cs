using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneManager
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private readonly ThisAddIn _addIn;
        private readonly ExcelInteropService _excelInteropService;
        private readonly ICaseTaskPaneSnapshotReader _caseTaskPaneSnapshotReader;
        private readonly TaskPaneBusinessActionLauncher _taskPaneBusinessActionLauncher;
        private readonly KernelCommandService _kernelCommandService;
        private readonly AccountingSheetCommandService _accountingSheetCommandService;
        private readonly CaseTaskPaneViewStateBuilder _caseTaskPaneViewStateBuilder;
        private readonly CasePaneSnapshotRenderService _casePaneSnapshotRenderService;
        private readonly AccountingInternalCommandService _accountingInternalCommandService;
        private readonly UserErrorService _userErrorService;
        private readonly KernelCaseInteractionState _kernelCaseInteractionState;
        private readonly Logger _logger;
        private readonly Dictionary<string, TaskPaneHost> _hostsByWindowKey;
        private readonly TaskPaneHostRegistry _taskPaneHostRegistry;
        private readonly TaskPaneDisplayCoordinator _taskPaneDisplayCoordinator;
        private readonly TaskPaneNonCaseActionHandler _taskPaneNonCaseActionHandler;
        private readonly TaskPaneActionDispatcher _taskPaneActionDispatcher;
        private readonly TaskPaneRefreshFlowCoordinator _taskPaneRefreshFlowCoordinator;
        private readonly CasePaneCacheRefreshNotificationService _casePaneCacheRefreshNotificationService;
        private readonly TaskPaneManagerTestHooks _testHooks;
        private int _kernelFlickerTraceRefreshPaneSequence;

        internal TaskPaneManager(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger)
            : this(
                addIn,
                excelInteropService,
                caseTaskPaneSnapshotReader,
                taskPaneBusinessActionLauncher,
                kernelCommandService,
                accountingSheetCommandService,
                caseTaskPaneViewStateBuilder,
                new CasePaneSnapshotRenderService(caseTaskPaneSnapshotReader, caseTaskPaneViewStateBuilder),
                accountingInternalCommandService,
                kernelCaseInteractionState,
                userErrorService,
                logger,
                testHooks: null)
        {
        }

        internal TaskPaneManager(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneManagerTestHooks testHooks)
            : this(
                addIn,
                excelInteropService,
                caseTaskPaneSnapshotReader,
                taskPaneBusinessActionLauncher,
                kernelCommandService,
                accountingSheetCommandService,
                caseTaskPaneViewStateBuilder,
                new CasePaneSnapshotRenderService(caseTaskPaneSnapshotReader, caseTaskPaneViewStateBuilder),
                accountingInternalCommandService,
                kernelCaseInteractionState,
                userErrorService,
                logger,
                testHooks)
        {
        }

        internal TaskPaneManager(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneManagerTestHooks testHooks)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _caseTaskPaneSnapshotReader = caseTaskPaneSnapshotReader ?? throw new ArgumentNullException(nameof(caseTaskPaneSnapshotReader));
            _taskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher ?? throw new ArgumentNullException(nameof(taskPaneBusinessActionLauncher));
            _kernelCommandService = kernelCommandService ?? throw new ArgumentNullException(nameof(kernelCommandService));
            _accountingSheetCommandService = accountingSheetCommandService ?? throw new ArgumentNullException(nameof(accountingSheetCommandService));
            _caseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            _casePaneSnapshotRenderService = casePaneSnapshotRenderService ?? throw new ArgumentNullException(nameof(casePaneSnapshotRenderService));
            _accountingInternalCommandService = accountingInternalCommandService ?? throw new ArgumentNullException(nameof(accountingInternalCommandService));
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _userErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
            _testHooks = testHooks;
            _casePaneCacheRefreshNotificationService = new CasePaneCacheRefreshNotificationService(
                _logger,
                workbook => _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook),
                _testHooks != null && _testHooks.OnCasePaneUpdatedNotification != null
                    ? new Action<string>(reason => _testHooks.OnCasePaneUpdatedNotification(reason))
                    : null);
            _taskPaneHostRegistry = new TaskPaneHostRegistry(
                _hostsByWindowKey,
                _addIn,
                _logger,
                FormatHostDescriptor,
                KernelControl_ActionInvoked,
                AccountingControl_ActionInvoked,
                (windowKey, control, e) => _taskPaneActionDispatcher?.HandleCaseControlActionInvoked(windowKey, control, e));
            _taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                _hostsByWindowKey,
                _kernelCaseInteractionState,
                _logger,
                _testHooks,
                SafeGetWindowKey,
                FormatHostDescriptor,
                workbook => FormatWorkbookDescriptor(workbook),
                FormatWindowDescriptor,
                windowKey => _taskPaneHostRegistry.RemoveHost(windowKey));
            _taskPaneNonCaseActionHandler = new TaskPaneNonCaseActionHandler(
                _excelInteropService,
                _kernelCommandService,
                _accountingSheetCommandService,
                _accountingInternalCommandService,
                _userErrorService,
                _logger,
                windowKey => _hostsByWindowKey.TryGetValue(windowKey ?? string.Empty, out TaskPaneHost host) ? host : null,
                RenderHost,
                (host, reason) => _taskPaneDisplayCoordinator.TryShowHost(host, reason));
            _taskPaneActionDispatcher = new TaskPaneActionDispatcher(
                _addIn,
                _excelInteropService,
                _taskPaneBusinessActionLauncher,
                _caseTaskPaneViewStateBuilder,
                _userErrorService,
                _logger,
                windowKey => _hostsByWindowKey.TryGetValue(windowKey ?? string.Empty, out TaskPaneHost host) ? host : null,
                host => _taskPaneDisplayCoordinator.InvalidateHostRenderStateForForcedRefresh(host),
                (control, workbook) => _casePaneSnapshotRenderService.RenderAfterAction(control, workbook),
                (host, reason) => _taskPaneDisplayCoordinator.TryShowHost(host, reason));
            _taskPaneRefreshFlowCoordinator = new TaskPaneRefreshFlowCoordinator(this);
        }

        internal TaskPaneManager(Logger logger, KernelCaseInteractionState kernelCaseInteractionState, TaskPaneManagerTestHooks testHooks)
        {
            _addIn = null;
            _excelInteropService = null;
            _caseTaskPaneSnapshotReader = null;
            _taskPaneBusinessActionLauncher = null;
            _kernelCommandService = null;
            _accountingSheetCommandService = null;
            _caseTaskPaneViewStateBuilder = null;
            _casePaneSnapshotRenderService = null;
            _accountingInternalCommandService = null;
            _userErrorService = null;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
            _testHooks = testHooks;
            _casePaneCacheRefreshNotificationService = new CasePaneCacheRefreshNotificationService(
                _logger,
                workbook => workbook == null ? string.Empty : (workbook.FullName ?? string.Empty),
                _testHooks != null && _testHooks.OnCasePaneUpdatedNotification != null
                    ? new Action<string>(reason => _testHooks.OnCasePaneUpdatedNotification(reason))
                    : null);
            _taskPaneHostRegistry = new TaskPaneHostRegistry(
                _hostsByWindowKey,
                _addIn,
                _logger,
                FormatHostDescriptor,
                KernelControl_ActionInvoked,
                AccountingControl_ActionInvoked,
                (windowKey, control, e) => _taskPaneActionDispatcher?.HandleCaseControlActionInvoked(windowKey, control, e));
            _taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                _hostsByWindowKey,
                _kernelCaseInteractionState,
                _logger,
                _testHooks,
                SafeGetWindowKey,
                FormatHostDescriptor,
                workbook => FormatWorkbookDescriptor(workbook),
                FormatWindowDescriptor,
                windowKey => _taskPaneHostRegistry.RemoveHost(windowKey));
            _taskPaneNonCaseActionHandler = null;
            _taskPaneActionDispatcher = null;
            _taskPaneRefreshFlowCoordinator = new TaskPaneRefreshFlowCoordinator(this);
        }

        internal bool RefreshPane(WorkbookContext context, string reason)
        {
            return _taskPaneRefreshFlowCoordinator.RefreshPane(context, reason);
        }

        internal bool TryShowExistingPane(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            return _taskPaneDisplayCoordinator.TryShowExistingPane(_excelInteropService, workbook, window, reason);
        }

        internal bool TryShowExistingPaneForDisplayRequest(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.TryShowExistingPaneForDisplayRequest(_excelInteropService, workbook, window);
        }

        internal bool ShouldShowWithRenderPaneForDisplayRequest(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.ShouldShowWithRenderPaneForDisplayRequest(_excelInteropService, workbook, window);
        }

        internal bool HasManagedPaneForWindow(Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.HasManagedPaneForWindow(window);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(_excelInteropService, workbook, window);
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
                HideAll();
                return false;
            }

            windowKey = SafeGetWindowKey(context.Window);
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
                HideAll();
                _logger.Warn("RefreshPane skipped because windowKey was empty. reason=" + (reason ?? string.Empty));
                return false;
            }

            return true;
        }

        private TaskPaneHost ResolveRefreshHost(WorkbookContext context, string windowKey, int refreshPaneCallId)
        {
            RemoveStaleKernelHosts(context, windowKey);
            TaskPaneHost host = GetOrReplaceHost(windowKey, context.Window, context.Role);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=host-selected refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", host="
                + FormatHostDescriptor(host));
            return host;
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
                + FormatHostDescriptor(host)
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

        private bool RenderAndShowHostForRefresh(TaskPaneHost host, WorkbookContext context, string reason, string windowKey, int refreshPaneCallId)
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
                + FormatHostDescriptor(host)
                + ", renderRequired="
                + renderState.IsRenderRequired.ToString());
            if (renderState.IsRenderRequired)
            {
                RenderHost(host, context, reason);
                host.LastRenderSignature = renderState.RenderSignature;
            }
            else
            {
                _logger.Debug(nameof(TaskPaneManager), "RefreshPane render skipped because the host state did not change. windowKey=" + windowKey + ", role=" + context.Role);
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
                + FormatHostDescriptor(host)
                + ", result=Shown");
            _logger.Info("TaskPane refreshed. reason=" + (reason ?? string.Empty) + ", role=" + context.Role + ", windowKey=" + windowKey);
            return true;
        }

        // Host lifecycle 責務: windowKey 単位の host 集合を保持し、hide/dispose/create/remove を担う。
        internal void HideAll()
        {
            _taskPaneDisplayCoordinator.HideAll();
        }

        internal void HideKernelPanes()
        {
            _taskPaneDisplayCoordinator.HideKernelPanes();
        }

        internal void HideAllExcept(string activeWindowKey)
        {
            _taskPaneDisplayCoordinator.HideAllExcept(activeWindowKey);
        }

        /// <summary>
        /// メソッド: host 表示前に、他ウィンドウ pane の表示状態を必要最小限だけ調整する。
        /// 引数: host - これから表示する host。
        /// 戻り値: なし。
        /// 副作用: CASE pane は複数窓で維持し、それ以外は従来どおり単一表示に寄せる。
        /// </summary>
        internal void PrepareHostsBeforeShow(TaskPaneHost host)
        {
            _taskPaneDisplayCoordinator.PrepareHostsBeforeShow(host);
        }

        /// <summary>
        /// メソッド: 指定 workbook に紐づく pane だけを非表示にして破棄する。
        /// 引数: workbook - 対象 workbook。
        /// 戻り値: なし。
        /// 副作用: 対象 workbook の Host を破棄し、他 workbook の pane には影響しない。
        /// </summary>
        internal void RemoveWorkbookPanes(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
            if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                return;
            }

            _taskPaneHostRegistry.RemoveWorkbookPanes(workbookFullName);
        }

        /// <summary>
        /// メソッド: 指定 window に紐づく pane だけを非表示にする。
        /// 引数: window - 対象 window。
        /// 戻り値: なし。
        /// 副作用: 対象 window の pane だけを非表示にする。
        /// </summary>
        internal void HidePaneForWindow(Excel.Window window)
        {
            _taskPaneDisplayCoordinator.HidePaneForWindow(window);
        }

        internal void DisposeAll()
        {
            _taskPaneHostRegistry.DisposeAll();
        }

        internal void RegisterHost(TaskPaneHost host)
        {
            _taskPaneHostRegistry.RegisterHost(host);
        }

        private TaskPaneHost GetOrReplaceHost(string windowKey, Excel.Window window, WorkbookRole role)
        {
            return _taskPaneHostRegistry.GetOrReplaceHost(windowKey, window, role);
        }

        private void RemoveStaleKernelHosts(WorkbookContext context, string activeWindowKey)
        {
            if (context == null
                || context.Role != WorkbookRole.Kernel
                || string.IsNullOrWhiteSpace(context.WorkbookFullName)
                || string.IsNullOrWhiteSpace(activeWindowKey))
            {
                return;
            }

            var staleKeys = new List<string>();
            foreach (KeyValuePair<string, TaskPaneHost> pair in _hostsByWindowKey)
            {
                if (string.Equals(pair.Key, activeWindowKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                TaskPaneHost host = pair.Value;
                if (!(host.Control is KernelNavigationControl))
                {
                    continue;
                }

                if (!string.Equals(host.WorkbookFullName, context.WorkbookFullName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                staleKeys.Add(pair.Key);
            }

            foreach (string staleKey in staleKeys)
            {
                _logger.Info(
                    "Removed stale kernel task pane host. workbook="
                    + context.WorkbookFullName
                    + ", staleWindowKey="
                    + staleKey
                    + ", activeWindowKey="
                    + activeWindowKey);
                RemoveHost(staleKey);
            }
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

        // Render 制御責務: role ごとの描画と signature 更新対象を分離し、再描画条件は上位から受け取る。
        private void RenderHost(TaskPaneHost host, WorkbookContext context, string reason)
        {
            host.WorkbookFullName = context.WorkbookFullName;

            if (host.Control is KernelNavigationControl kernelControl)
            {
                RenderKernelHost(kernelControl, context);
                return;
            }

            if (host.Control is AccountingNavigationControl accountingControl)
            {
                RenderAccountingHost(accountingControl, context);
                return;
            }

            if (host.Control is DocumentButtonsControl caseControl)
            {
                RenderCaseHost(caseControl, context, reason);
            }
        }

        private void RenderKernelHost(KernelNavigationControl kernelControl, WorkbookContext context)
        {
            _logger.Info("RenderHost start. role=Kernel, workbook=" + (context.WorkbookFullName ?? string.Empty));
            kernelControl.Render(KernelNavigationDefinitions.CreateForSheet(context.ActiveSheetCodeName));
            _logger.Info("RenderHost completed. role=Kernel, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        private void RenderAccountingHost(AccountingNavigationControl accountingControl, WorkbookContext context)
        {
            _logger.Info("RenderHost start. role=Accounting, workbook=" + (context.WorkbookFullName ?? string.Empty));
            accountingControl.Render(AccountingNavigationDefinitions.CreateForSheet(context.ActiveSheetCodeName));
            _logger.Info("RenderHost completed. role=Accounting, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        private void RenderCaseHost(DocumentButtonsControl caseControl, WorkbookContext context, string reason)
        {
            _logger.Info("RenderHost start. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
            bool? originalWorkbookSavedState = _casePaneCacheRefreshNotificationService.TryGetWorkbookSavedState(context.Workbook);
            CasePaneSnapshotRenderService.CasePaneSnapshotRenderResult renderResult = _casePaneSnapshotRenderService.Render(caseControl, context.Workbook);
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult = renderResult.BuildResult;
            string snapshotText = buildResult.SnapshotText;
            _logger.Info("RenderHost snapshot acquired. role=Case, length=" + snapshotText.Length.ToString());
            TaskPaneSnapshot snapshot = renderResult.Snapshot;
            _logger.Info("RenderHost snapshot parsed. role=Case, hasError=" + snapshot.HasError.ToString() + ", tabs=" + snapshot.Tabs.Count.ToString() + ", docs=" + snapshot.DocButtons.Count.ToString());
            _casePaneCacheRefreshNotificationService.NotifyCasePaneUpdatedIfNeeded(context.Workbook, reason, buildResult, originalWorkbookSavedState);
            _logger.Info("RenderHost completed. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        internal void PrepareTargetWindowForForcedRefresh(Excel.Window targetWindow)
        {
            _taskPaneDisplayCoordinator.PrepareTargetWindowForForcedRefresh(targetWindow);
        }

        private bool TryShowHost(TaskPaneHost host, string reason)
        {
            return _taskPaneDisplayCoordinator.TryShowHost(host, reason);
        }

        private void RemoveHost(string windowKey)
        {
            _taskPaneHostRegistry.RemoveHost(windowKey);
        }

        private void KernelControl_ActionInvoked(string windowKey, KernelNavigationActionEventArgs e)
        {
            _taskPaneNonCaseActionHandler?.HandleKernelActionInvoked(windowKey, e);
        }

        /// <summary>
        /// メソッド: 会計 pane のボタン押下を受けて内部処理を実行する。
        /// 引数: windowKey - 対象 host の window key, e - アクション引数。
        /// 戻り値: なし。
        /// 副作用: 会計ブック内部処理を実行し、必要に応じて pane を再描画する。
        /// </summary>
        private void AccountingControl_ActionInvoked(string windowKey, AccountingNavigationActionEventArgs e)
        {
            _taskPaneNonCaseActionHandler?.HandleAccountingActionInvoked(windowKey, e);
        }

        private static string SafeGetWindowKey(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd) ?? string.Empty;
            }
            catch
            {
                // window key を取得できない場合は空文字へフォールバックする。
                return string.Empty;
            }
        }

        private string FormatContextDescriptor(WorkbookContext context)
        {
            if (context == null)
            {
                return "null";
            }

            return "role=\""
                + context.Role.ToString()
                + "\",workbook="
                + FormatWorkbookDescriptor(context.Workbook, context.WorkbookFullName)
                + ",window="
                + FormatWindowDescriptor(context.Window)
                + ",activeSheet=\""
                + (context.ActiveSheetCodeName ?? string.Empty)
                + "\"";
        }

        private string FormatHostDescriptor(TaskPaneHost host)
        {
            if (host == null)
            {
                return "null";
            }

            return "paneRole=\""
                + GetPaneRoleName(host)
                + "\",windowKey=\""
                + (host.WindowKey ?? string.Empty)
                + "\",workbookFullName=\""
                + (host.WorkbookFullName ?? string.Empty)
                + "\",window="
                + FormatWindowDescriptor(host.Window);
        }

        private string GetPaneRoleName(TaskPaneHost host)
        {
            if (host == null || host.Control == null)
            {
                return "Unknown";
            }

            if (host.Control is DocumentButtonsControl)
            {
                return "Case";
            }

            if (host.Control is KernelNavigationControl)
            {
                return "Kernel";
            }

            if (host.Control is AccountingNavigationControl)
            {
                return "Accounting";
            }

            return host.Control.GetType().Name;
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return FormatWorkbookDescriptor(workbook, null);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook, string fallbackFullName)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook, fallbackFullName)
                + "\",name=\""
                + SafeWorkbookShortName(workbook)
                + "\"";
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook, string fallbackFullName)
        {
            string workbookFullName = workbook == null || _excelInteropService == null
                ? string.Empty
                : (_excelInteropService.GetWorkbookFullName(workbook) ?? string.Empty);
            return string.IsNullOrWhiteSpace(workbookFullName) ? (fallbackFullName ?? string.Empty) : workbookFullName;
        }

        private string SafeWorkbookShortName(Excel.Workbook workbook)
        {
            return workbook == null || _excelInteropService == null
                ? string.Empty
                : (_excelInteropService.GetWorkbookName(workbook) ?? string.Empty);
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeGetWindowKey(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\"";
        }

        private static string SafeWindowCaption(Excel.Window window)
        {
            try
            {
                if (window == null)
                {
                    return string.Empty;
                }

                dynamic lateBoundWindow = window;
                return Convert.ToString(lateBoundWindow.Caption) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal sealed class TaskPaneManagerTestHooks
        {
            internal Action<string, string> OnHideHost { get; set; }

            internal Func<string, string, bool> TryShowHost { get; set; }

            internal Action<string> OnCasePaneUpdatedNotification { get; set; }
        }

        // Refresh flow 責務: RefreshPane の主経路で前提確認、host 解決、reuse、render/show を順序どおり調停する。
        private sealed class TaskPaneRefreshFlowCoordinator
        {
            private readonly TaskPaneManager _owner;

            internal TaskPaneRefreshFlowCoordinator(TaskPaneManager owner)
            {
                _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            }

            internal bool RefreshPane(WorkbookContext context, string reason)
            {
                int refreshPaneCallId = ++_owner._kernelFlickerTraceRefreshPaneSequence;
                _owner._logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=refresh-pane-start refreshPaneCallId="
                    + refreshPaneCallId.ToString()
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", context="
                    + _owner.FormatContextDescriptor(context));
                if (!_owner.TryAcceptRefreshPaneRequest(context, reason, refreshPaneCallId, out WorkbookRole role, out string windowKey))
                {
                    return false;
                }

                TaskPaneHost host = _owner.ResolveRefreshHost(context, windowKey, refreshPaneCallId);
                if (_owner.TryReuseCaseHostForRefresh(host, context, reason, windowKey, refreshPaneCallId))
                {
                    return true;
                }

                return _owner.RenderAndShowHostForRefresh(host, context, reason, windowKey, refreshPaneCallId);
            }
        }
    }

    internal static class TaskPaneShowExistingPolicy
    {
        internal static bool ShouldShowExisting(bool hasExistingHost, bool isSameWorkbook, bool isRenderSignatureCurrent)
        {
            return hasExistingHost && isSameWorkbook && isRenderSignatureCurrent;
        }
    }

    internal static class TaskPaneShowWithRenderPolicy
    {
        internal static bool ShouldShowWithRender(bool hasExistingHost, bool isSameWorkbook, bool isRenderSignatureCurrent)
        {
            return !hasExistingHost || !isSameWorkbook || !isRenderSignatureCurrent;
        }
    }
}
