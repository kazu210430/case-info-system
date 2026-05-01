using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

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
        private readonly TaskPaneActionDispatcher _taskPaneActionDispatcher;
        private readonly TaskPaneRefreshFlowCoordinator _taskPaneRefreshFlowCoordinator;
        private readonly TaskPaneManagerTestHooks _testHooks;
        private const string ProductTitle = "案件情報System";
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
            _taskPaneHostRegistry = new TaskPaneHostRegistry(
                _hostsByWindowKey,
                _addIn,
                _logger,
                FormatHostDescriptor,
                KernelControl_ActionInvoked,
                AccountingControl_ActionInvoked,
                (windowKey, control, e) => _taskPaneActionDispatcher.HandleCaseControlActionInvoked(windowKey, control, e));
            _taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                _hostsByWindowKey,
                _kernelCaseInteractionState,
                _logger,
                _testHooks,
                SafeGetWindowKey,
                FormatHostDescriptor,
                workbook => FormatWorkbookDescriptor(workbook),
                windowKey => _taskPaneHostRegistry.RemoveHost(windowKey));
            _taskPaneActionDispatcher = new TaskPaneActionDispatcher(this);
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
            _taskPaneHostRegistry = new TaskPaneHostRegistry(
                _hostsByWindowKey,
                _addIn,
                _logger,
                FormatHostDescriptor,
                KernelControl_ActionInvoked,
                AccountingControl_ActionInvoked,
                (windowKey, control, e) => _taskPaneActionDispatcher.HandleCaseControlActionInvoked(windowKey, control, e));
            _taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                _hostsByWindowKey,
                _kernelCaseInteractionState,
                _logger,
                _testHooks,
                SafeGetWindowKey,
                FormatHostDescriptor,
                workbook => FormatWorkbookDescriptor(workbook),
                windowKey => _taskPaneHostRegistry.RemoveHost(windowKey));
            _taskPaneActionDispatcher = new TaskPaneActionDispatcher(this);
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
            TaskPaneDisplayRequestPaneState paneState = TaskPaneRenderStateEvaluator.EvaluateDisplayRequestPaneState(
                _excelInteropService,
                _hostsByWindowKey,
                workbook,
                window);

            if (!TaskPaneShowExistingPolicy.ShouldShowExisting(
                hasExistingHost: paneState.HasExistingHost,
                isSameWorkbook: paneState.IsSameWorkbook,
                isRenderSignatureCurrent: paneState.IsRenderSignatureCurrent))
            {
                return false;
            }

            return TryShowExistingPane(workbook, window, "DisplayRequest.ShowExisting");
        }

        internal bool ShouldShowWithRenderPaneForDisplayRequest(Excel.Workbook workbook, Excel.Window window)
        {
            TaskPaneDisplayRequestPaneState paneState = TaskPaneRenderStateEvaluator.EvaluateDisplayRequestPaneState(
                _excelInteropService,
                _hostsByWindowKey,
                workbook,
                window);

            return TaskPaneShowWithRenderPolicy.ShouldShowWithRender(
                paneState.HasExistingHost,
                paneState.IsSameWorkbook,
                paneState.IsRenderSignatureCurrent);
        }

        internal bool HasManagedPaneForWindow(Excel.Window window)
        {
            string windowKey = SafeGetWindowKey(window);
            return !string.IsNullOrWhiteSpace(windowKey)
                && _hostsByWindowKey.ContainsKey(windowKey);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            string windowKey = SafeGetWindowKey(window);
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=visible-case-pane-check result=NoWindowKey"
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", inputWindow="
                    + FormatWindowDescriptor(window));
                return false;
            }

            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=visible-case-pane-check result=NoHost"
                    + ", windowKey="
                    + windowKey
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook));
                return false;
            }

            string workbookFullName = workbook == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
            if (string.IsNullOrWhiteSpace(workbookFullName)
                || !string.Equals(host.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=visible-case-pane-check result=WorkbookMismatch"
                    + ", windowKey="
                    + windowKey
                    + ", host="
                    + FormatHostDescriptor(host)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook));
                return false;
            }

            WorkbookRole hostedRole = GetHostedWorkbookRole(host);
            bool isVisibleCasePane = hostedRole == WorkbookRole.Case && host.IsVisible;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=visible-case-pane-check result="
                + (isVisibleCasePane ? "VisibleCasePaneFound" : "NotVisibleOrNotCase")
                + ", windowKey="
                + windowKey
                + ", host="
                + FormatHostDescriptor(host)
                + ", hostedRole="
                + hostedRole.ToString()
                + ", hostVisible="
                + host.IsVisible.ToString());
            return isVisibleCasePane;
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
            bool? originalWorkbookSavedState = TryGetWorkbookSavedState(context.Workbook);
            CasePaneSnapshotRenderService.CasePaneSnapshotRenderResult renderResult = _casePaneSnapshotRenderService.Render(caseControl, context.Workbook);
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult = renderResult.BuildResult;
            string snapshotText = buildResult.SnapshotText;
            _logger.Info("RenderHost snapshot acquired. role=Case, length=" + snapshotText.Length.ToString());
            TaskPaneSnapshot snapshot = renderResult.Snapshot;
            _logger.Info("RenderHost snapshot parsed. role=Case, hasError=" + snapshot.HasError.ToString() + ", tabs=" + snapshot.Tabs.Count.ToString() + ", docs=" + snapshot.DocButtons.Count.ToString());
            NotifyCasePaneUpdatedIfNeeded(context.Workbook, reason, buildResult, originalWorkbookSavedState);
            _logger.Info("RenderHost completed. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        /// <summary>
        /// メソッド: CASE の文書ボタンパネル内部更新後に、開く導線専用の後処理を適用する。
        /// 引数: workbook - 対象 CASE ブック, reason - pane 更新理由, buildResult - スナップショット生成結果。
        /// 戻り値: なし。
        /// 副作用: 内部キャッシュ更新による保存確認を抑止し、業務メッセージを表示する。
        /// </summary>
        internal void NotifyCasePaneUpdatedIfNeeded(Excel.Workbook workbook, string reason, TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult, bool? originalSavedState = null)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                bool updatedCaseSnapshotCache = buildResult != null && buildResult.UpdatedCaseSnapshotCache;
                if (updatedCaseSnapshotCache)
                {
                    RestoreWorkbookSavedState(workbook, originalSavedState);
                }

                if (!CasePaneCacheRefreshNotificationPolicy.ShouldNotify(updatedCaseSnapshotCache, reason))
                {
                    return;
                }

                if (_testHooks != null && _testHooks.OnCasePaneUpdatedNotification != null)
                {
                    _testHooks.OnCasePaneUpdatedNotification(reason ?? string.Empty);
                    return;
                }

                MessageBox.Show("文書ボタンパネルを更新しました", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                _logger.Info("CASE pane cache refresh notification was shown. workbook=" + SafeGetWorkbookName(workbook) + ", reason=" + (reason ?? string.Empty));
            }
            catch (Exception ex)
            {
                // 例外処理: 通知失敗で CASE 表示自体を止めないため、ログのみ残して継続する。
                _logger.Error("NotifyCasePaneUpdatedIfNeeded failed.", ex);
            }
        }

        /// <summary>
        /// メソッド: ログ出力用に workbook 名を安全に取得する。
        /// 引数: workbook - 対象 workbook。
        /// 戻り値: workbook フルパス。取得できない場合は空文字。
        /// 副作用: なし。
        /// </summary>
        private string SafeGetWorkbookName(Excel.Workbook workbook)
        {
            return workbook == null ? string.Empty : (_excelInteropService.GetWorkbookFullName(workbook) ?? string.Empty);
        }

        private bool? TryGetWorkbookSavedState(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            try
            {
                return workbook.Saved;
            }
            catch (Exception ex)
            {
                _logger.Error("TryGetWorkbookSavedState failed.", ex);
                return null;
            }
        }

        private void RestoreWorkbookSavedState(Excel.Workbook workbook, bool? originalSavedState)
        {
            if (workbook == null || !originalSavedState.HasValue)
            {
                return;
            }

            try
            {
                workbook.Saved = originalSavedState.Value;
            }
            catch (Exception ex)
            {
                _logger.Error("RestoreWorkbookSavedState failed.", ex);
            }
        }

        internal void PrepareTargetWindowForForcedRefresh(Excel.Window targetWindow)
        {
            _taskPaneDisplayCoordinator.PrepareTargetWindowForForcedRefresh(targetWindow);
        }

        private void InvalidateHostRenderStateForForcedRefresh(TaskPaneHost host)
        {
            _taskPaneDisplayCoordinator.InvalidateHostRenderStateForForcedRefresh(host);
        }

        /// <summary>
        /// メソッド: CASE pane アクション後の再描画準備と描画を行う。
        /// 引数: control - 描画対象 control, workbook - 対象 workbook。
        /// 戻り値: なし。
        /// 副作用: 最新 snapshot から ViewState を再構築し、選択タブを維持したまま CASE pane を再描画する。
        /// </summary>
        private void RenderCaseHostAfterAction(DocumentButtonsControl control, Excel.Workbook workbook)
        {
            _casePaneSnapshotRenderService.RenderAfterAction(control, workbook);
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
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                _logger.Warn("KernelControl_ActionInvoked skipped because windowKey was empty.");
                return;
            }

            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                _logger.Warn("KernelControl_ActionInvoked skipped because host was not found. windowKey=" + windowKey);
                return;
            }

            Excel.Workbook workbook = _excelInteropService.FindOpenWorkbook(host.WorkbookFullName);
            if (workbook == null)
            {
                _logger.Warn("KernelControl_ActionInvoked skipped because workbook was not found. windowKey=" + windowKey);
                return;
            }

            WorkbookContext context = BuildWorkbookContext(host, workbook, WorkbookRole.Kernel, _excelInteropService.GetActiveSheetCodeName(workbook));
            _kernelCommandService.Execute(context, e.ActionId);
        }

        /// <summary>
        /// メソッド: 会計 pane のボタン押下を受けて内部処理を実行する。
        /// 引数: windowKey - 対象 host の window key, e - アクション引数。
        /// 戻り値: なし。
        /// 副作用: 会計ブック内部処理を実行し、必要に応じて pane を再描画する。
        /// </summary>
        private void AccountingControl_ActionInvoked(string windowKey, AccountingNavigationActionEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                _logger.Warn("AccountingControl_ActionInvoked skipped because windowKey was empty.");
                return;
            }

            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                _logger.Warn("AccountingControl_ActionInvoked skipped because host was not found. windowKey=" + windowKey);
                return;
            }

            Excel.Workbook workbook = _excelInteropService.FindOpenWorkbook(host.WorkbookFullName);
            if (workbook == null)
            {
                _logger.Warn("AccountingControl_ActionInvoked skipped because workbook was not found. windowKey=" + windowKey);
                return;
            }

            try
            {
                WorkbookContext context = BuildWorkbookContext(host, workbook, WorkbookRole.Accounting, TryGetWorksheetCodeName(workbook));
                _accountingSheetCommandService.Execute(context, e.ActionId);
                _accountingInternalCommandService.Execute(context, e.ActionId);

                WorkbookContext refreshedContext = BuildWorkbookContext(host, workbook, WorkbookRole.Accounting, _excelInteropService.GetActiveSheetCodeName(workbook));
                RenderHost(host, refreshedContext, "AccountingControl_ActionInvoked");
                TryShowHost(host, "AccountingControl_ActionInvoked");
            }
            catch (Exception ex)
            {
                _logger.Error("AccountingControl_ActionInvoked failed.", ex);
                _userErrorService.ShowUserError("AccountingControl_ActionInvoked", ex);
            }
        }

        private WorkbookContext BuildWorkbookContext(TaskPaneHost host, Excel.Workbook workbook, WorkbookRole role, string activeSheetCodeName)
        {
            return new WorkbookContext(
                workbook,
                host.Window,
                role,
                _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT"),
                _excelInteropService.GetWorkbookFullName(workbook),
                activeSheetCodeName);
        }

        private static string TryGetWorksheetCodeName(Excel.Workbook workbook)
        {
            try
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;
                return worksheet == null ? string.Empty : worksheet.CodeName ?? string.Empty;
            }
            catch
            {
                // 会計 pane の再描画判定で CodeName を取得できない場合は安全側で空文字へフォールバックする。
                return string.Empty;
            }
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

        private static WorkbookRole GetHostedWorkbookRole(TaskPaneHost host)
        {
            if (host == null || host.Control == null)
            {
                return WorkbookRole.Unknown;
            }

            if (host.Control is DocumentButtonsControl)
            {
                return WorkbookRole.Case;
            }

            if (host.Control is KernelNavigationControl)
            {
                return WorkbookRole.Kernel;
            }

            if (host.Control is AccountingNavigationControl)
            {
                return WorkbookRole.Accounting;
            }

            return WorkbookRole.Unknown;
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

        private sealed class TaskPaneHostRegistry
        {
            private readonly Dictionary<string, TaskPaneHost> _hostsByWindowKey;
            private readonly ThisAddIn _addIn;
            private readonly Logger _logger;
            private readonly Func<TaskPaneHost, string> _formatHostDescriptor;
            private readonly Action<string, KernelNavigationActionEventArgs> _handleKernelActionInvoked;
            private readonly Action<string, AccountingNavigationActionEventArgs> _handleAccountingActionInvoked;
            private readonly Action<string, DocumentButtonsControl, TaskPaneActionEventArgs> _handleCaseActionInvoked;

            internal TaskPaneHostRegistry(
                Dictionary<string, TaskPaneHost> hostsByWindowKey,
                ThisAddIn addIn,
                Logger logger,
                Func<TaskPaneHost, string> formatHostDescriptor,
                Action<string, KernelNavigationActionEventArgs> handleKernelActionInvoked,
                Action<string, AccountingNavigationActionEventArgs> handleAccountingActionInvoked,
                Action<string, DocumentButtonsControl, TaskPaneActionEventArgs> handleCaseActionInvoked)
            {
                _hostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
                _addIn = addIn;
                _logger = logger ?? throw new ArgumentNullException(nameof(logger));
                _formatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
                _handleKernelActionInvoked = handleKernelActionInvoked ?? throw new ArgumentNullException(nameof(handleKernelActionInvoked));
                _handleAccountingActionInvoked = handleAccountingActionInvoked ?? throw new ArgumentNullException(nameof(handleAccountingActionInvoked));
                _handleCaseActionInvoked = handleCaseActionInvoked ?? throw new ArgumentNullException(nameof(handleCaseActionInvoked));
            }

            internal void RegisterHost(TaskPaneHost host)
            {
                if (host == null)
                {
                    throw new ArgumentNullException(nameof(host));
                }

                if (_hostsByWindowKey.TryGetValue(host.WindowKey, out TaskPaneHost existingHost)
                    && !ReferenceEquals(existingHost, host))
                {
                    existingHost.Dispose();
                }

                _hostsByWindowKey[host.WindowKey] = host;
            }

            internal TaskPaneHost GetOrReplaceHost(string windowKey, Excel.Window window, WorkbookRole role)
            {
                if (_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost existingHost))
                {
                    bool roleMatches =
                        (role == WorkbookRole.Kernel && existingHost.Control is KernelNavigationControl)
                        || (role == WorkbookRole.Case && existingHost.Control is DocumentButtonsControl)
                        || (role == WorkbookRole.Accounting && existingHost.Control is AccountingNavigationControl);
                    if (roleMatches)
                    {
                        return existingHost;
                    }

                    existingHost.Dispose();
                    _hostsByWindowKey.Remove(windowKey);
                }

                if (role == WorkbookRole.Kernel)
                {
                    var kernelControl = new KernelNavigationControl();
                    kernelControl.ActionInvoked += (sender, e) => _handleKernelActionInvoked(windowKey, e);
                    var host = new TaskPaneHost(_addIn, window, kernelControl, kernelControl, windowKey);
                    _hostsByWindowKey.Add(windowKey, host);
                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneManager action=create-host host="
                        + _formatHostDescriptor(host)
                        + ", paneRole=Kernel");
                    _logger.Info("TaskPane host created. role=Kernel, windowKey=" + windowKey);
                    return host;
                }

                if (role == WorkbookRole.Accounting)
                {
                    var accountingControl = new AccountingNavigationControl();
                    accountingControl.ActionInvoked += (sender, e) => _handleAccountingActionInvoked(windowKey, e);
                    var host = new TaskPaneHost(_addIn, window, accountingControl, accountingControl, windowKey);
                    _hostsByWindowKey.Add(windowKey, host);
                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneManager action=create-host host="
                        + _formatHostDescriptor(host)
                        + ", paneRole=Accounting");
                    _logger.Info("TaskPane host created. role=Accounting, windowKey=" + windowKey);
                    return host;
                }

                var caseControl = new DocumentButtonsControl();
                var caseHost = new TaskPaneHost(_addIn, window, caseControl, caseControl, windowKey);
                caseControl.ActionInvoked += (sender, e) => _handleCaseActionInvoked(windowKey, caseControl, e);
                _hostsByWindowKey.Add(windowKey, caseHost);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneManager action=create-host host="
                    + _formatHostDescriptor(caseHost)
                    + ", paneRole=Case");
                _logger.Info("TaskPane host created. role=Case, windowKey=" + windowKey);
                return caseHost;
            }

            internal void RemoveWorkbookPanes(string workbookFullName)
            {
                var targetKeys = new List<string>();
                foreach (KeyValuePair<string, TaskPaneHost> pair in _hostsByWindowKey)
                {
                    if (string.Equals(pair.Value.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetKeys.Add(pair.Key);
                    }
                }

                foreach (string windowKey in targetKeys)
                {
                    RemoveHost(windowKey);
                }
            }

            internal void DisposeAll()
            {
                foreach (TaskPaneHost host in new List<TaskPaneHost>(_hostsByWindowKey.Values))
                {
                    host.Dispose();
                }

                _hostsByWindowKey.Clear();
            }

            internal void RemoveHost(string windowKey)
            {
                if (string.IsNullOrWhiteSpace(windowKey))
                {
                    return;
                }

                if (_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
                {
                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneManager action=remove-host host="
                        + _formatHostDescriptor(host));
                    _hostsByWindowKey.Remove(windowKey);
                    try
                    {
                        host.Dispose();
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("TaskPane host dispose failed. windowKey=" + windowKey, ex);
                    }
                }
            }
        }

        private sealed class TaskPaneActionDispatcher
        {
            private readonly TaskPaneManager _owner;

            internal TaskPaneActionDispatcher(TaskPaneManager owner)
            {
                _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            }

            internal void HandleCaseControlActionInvoked(string windowKey, DocumentButtonsControl control, TaskPaneActionEventArgs e)
            {
                if (string.IsNullOrWhiteSpace(windowKey) || control == null)
                {
                    _owner._logger.Warn("CaseControl_ActionInvoked skipped because host identity was not available.");
                    return;
                }

                if (!_owner._hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
                {
                    _owner._logger.Warn("CaseControl_ActionInvoked skipped because host was not found. windowKey=" + windowKey);
                    return;
                }

                Excel.Workbook workbook = _owner._excelInteropService.FindOpenWorkbook(host.WorkbookFullName);
                if (workbook == null)
                {
                    _owner._logger.Warn("CaseControl_ActionInvoked skipped because workbook was not found. windowKey=" + windowKey);
                    control.Render(_owner._caseTaskPaneViewStateBuilder.BuildWorkbookNotFoundState());
                    return;
                }

                try
                {
                    bool shouldContinue = _owner._taskPaneBusinessActionLauncher.TryExecute(workbook, e.ActionKind, e.Key);
                    if (!shouldContinue)
                    {
                        return;
                    }

                    HandlePostActionRefresh(host, workbook, control, e.ActionKind);
                }
                catch (Exception ex)
                {
                    _owner._logger.Error("CaseControl_ActionInvoked failed.", ex);
                    control.Render(_owner._caseTaskPaneViewStateBuilder.BuildActionFailedState());
                    _owner._userErrorService.ShowUserError("CaseControl_ActionInvoked", ex);
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
                    _owner._logger.Info("CASE pane refresh after action skipped because " + reason);
                    return;
                }

                if (decision == TaskPanePostActionRefreshDecision.DeferAndInvalidateSignature)
                {
                    _owner.InvalidateHostRenderStateForForcedRefresh(host);
                    _owner._logger.Info("CASE pane refresh after case-list action was deferred so Kernel navigation can take the foreground.");
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

                if (_owner._addIn != null && host.Window != null)
                {
                    _owner._addIn.RequestTaskPaneDisplayForTargetWindow(
                        TaskPaneDisplayRequest.ForPostActionRefresh(actionKind),
                        workbook,
                        host.Window);
                    return;
                }

                _owner.InvalidateHostRenderStateForForcedRefresh(host);
                _owner.RenderCaseHostAfterAction(control, workbook);
                host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(
                    _owner._excelInteropService,
                    new WorkbookContext(
                        workbook,
                        host.Window,
                        WorkbookRole.Case,
                        _owner._excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT"),
                        _owner._excelInteropService.GetWorkbookFullName(workbook),
                        _owner._excelInteropService.GetActiveSheetCodeName(workbook)));
                if (!_owner.TryShowHost(host, "RefreshCaseHostAfterAction"))
                {
                    _owner._logger.Warn("CASE pane refresh after action skipped because host could not be shown. workbook=" + (host.WorkbookFullName ?? string.Empty));
                    return;
                }

                _owner._logger.Info("CASE pane refreshed after action. workbook=" + (host.WorkbookFullName ?? string.Empty));
            }
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
