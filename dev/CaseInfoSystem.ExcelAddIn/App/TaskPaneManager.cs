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
        private readonly TaskPaneHostLifecycleService _taskPaneHostLifecycleService;
        private readonly TaskPaneDisplayCoordinator _taskPaneDisplayCoordinator;
        private readonly TaskPaneNonCaseActionHandler _taskPaneNonCaseActionHandler;
        private readonly TaskPaneActionDispatcher _taskPaneActionDispatcher;
        private readonly TaskPaneHostFlowService _taskPaneHostFlowService;
        private readonly CasePaneCacheRefreshNotificationService _casePaneCacheRefreshNotificationService;
        private readonly TaskPaneManagerTestHooks _testHooks;

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
            _taskPaneHostLifecycleService = new TaskPaneHostLifecycleService(
                _hostsByWindowKey,
                _taskPaneHostRegistry,
                _excelInteropService,
                _logger);
            _taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                _hostsByWindowKey,
                _kernelCaseInteractionState,
                _logger,
                _testHooks,
                TaskPaneManagerDiagnosticHelper.SafeGetWindowKey,
                FormatHostDescriptor,
                workbook => FormatWorkbookDescriptor(workbook),
                TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor,
                windowKey => _taskPaneHostLifecycleService.RemoveHost(windowKey));
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
            Func<string, TaskPaneHost> resolveHost = windowKey => _hostsByWindowKey.TryGetValue(windowKey ?? string.Empty, out TaskPaneHost host) ? host : null;
            var taskPaneCaseFallbackActionExecutor = new TaskPaneCaseFallbackActionExecutor(_taskPaneBusinessActionLauncher);
            var taskPaneCaseActionTargetResolver = new TaskPaneCaseActionTargetResolver(
                _excelInteropService,
                _logger,
                resolveHost);
            Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh =
                (host, workbook, control, actionKind) => _taskPaneActionDispatcher.HandlePostActionRefresh(host, workbook, control, actionKind);
            var taskPaneCaseAccountingActionHandler = new TaskPaneCaseAccountingActionHandler(
                taskPaneCaseActionTargetResolver,
                taskPaneCaseFallbackActionExecutor,
                _caseTaskPaneViewStateBuilder,
                _userErrorService,
                _logger,
                handlePostActionRefresh);
            var taskPaneCaseDocumentActionHandler = new TaskPaneCaseDocumentActionHandler(
                taskPaneCaseActionTargetResolver,
                taskPaneCaseFallbackActionExecutor,
                _caseTaskPaneViewStateBuilder,
                _userErrorService,
                _logger,
                handlePostActionRefresh);
            _taskPaneActionDispatcher = new TaskPaneActionDispatcher(
                _addIn,
                _excelInteropService,
                _caseTaskPaneViewStateBuilder,
                _userErrorService,
                _logger,
                taskPaneCaseFallbackActionExecutor,
                taskPaneCaseActionTargetResolver,
                taskPaneCaseAccountingActionHandler,
                taskPaneCaseDocumentActionHandler,
                host => _taskPaneDisplayCoordinator.InvalidateHostRenderStateForForcedRefresh(host),
                (control, workbook) => _casePaneSnapshotRenderService.RenderAfterAction(control, workbook),
                (host, reason) => _taskPaneDisplayCoordinator.TryShowHost(host, reason));
            _taskPaneHostFlowService = new TaskPaneHostFlowService(
                _excelInteropService,
                _taskPaneDisplayCoordinator,
                _taskPaneHostLifecycleService,
                _logger,
                FormatContextDescriptor,
                FormatHostDescriptor,
                TaskPaneManagerDiagnosticHelper.SafeGetWindowKey,
                RenderHost);
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
            _taskPaneHostLifecycleService = new TaskPaneHostLifecycleService(
                _hostsByWindowKey,
                _taskPaneHostRegistry,
                _excelInteropService,
                _logger);
            _taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                _hostsByWindowKey,
                _kernelCaseInteractionState,
                _logger,
                _testHooks,
                TaskPaneManagerDiagnosticHelper.SafeGetWindowKey,
                FormatHostDescriptor,
                workbook => FormatWorkbookDescriptor(workbook),
                TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor,
                windowKey => _taskPaneHostLifecycleService.RemoveHost(windowKey));
            _taskPaneNonCaseActionHandler = null;
            _taskPaneActionDispatcher = null;
            _taskPaneHostFlowService = new TaskPaneHostFlowService(
                _excelInteropService,
                _taskPaneDisplayCoordinator,
                _taskPaneHostLifecycleService,
                _logger,
                FormatContextDescriptor,
                FormatHostDescriptor,
                TaskPaneManagerDiagnosticHelper.SafeGetWindowKey,
                RenderHost);
        }

        internal bool RefreshPane(WorkbookContext context, string reason)
        {
            return _taskPaneHostFlowService.RefreshPane(context, reason);
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

        // Display 制御 facade: 表示前調停と hide/show は display coordinator に委譲する。
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

        // Host lifecycle facade: workbook / host 単位の register/remove/dispose は lifecycle service に委譲する。
        /// <summary>
        /// メソッド: 指定 workbook に紐づく pane だけを非表示にして破棄する。
        /// 引数: workbook - 対象 workbook。
        /// 戻り値: なし。
        /// 副作用: 対象 workbook の Host を破棄し、他 workbook の pane には影響しない。
        /// </summary>
        internal void RemoveWorkbookPanes(Excel.Workbook workbook)
        {
            _taskPaneHostLifecycleService.RemoveWorkbookPanes(workbook);
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
            _taskPaneHostLifecycleService.DisposeAll();
        }

        internal void RegisterHost(TaskPaneHost host)
        {
            _taskPaneHostLifecycleService.RegisterHost(host);
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

        private string FormatContextDescriptor(WorkbookContext context)
        {
            return TaskPaneManagerDiagnosticHelper.FormatContextDescriptor(
                context,
                context == null ? string.Empty : FormatWorkbookDescriptor(context.Workbook, context.WorkbookFullName),
                context == null ? string.Empty : TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor(context.Window));
        }

        private string FormatHostDescriptor(TaskPaneHost host)
        {
            return TaskPaneManagerDiagnosticHelper.FormatHostDescriptor(
                host,
                TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor(host == null ? null : host.Window));
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

        private static class TaskPaneManagerDiagnosticHelper
        {
            internal static bool IsCaseHost(TaskPaneHost host)
            {
                return host != null && host.Control is DocumentButtonsControl;
            }

            internal static bool IsKernelHost(TaskPaneHost host)
            {
                return host != null && host.Control is KernelNavigationControl;
            }

            internal static string SafeGetWindowKey(Excel.Window window)
            {
                try
                {
                    return window == null ? string.Empty : Convert.ToString(window.Hwnd) ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }

            internal static string FormatContextDescriptor(WorkbookContext context, string workbookDescriptor, string windowDescriptor)
            {
                if (context == null)
                {
                    return "null";
                }

                return "role=\""
                    + context.Role.ToString()
                    + "\",workbook="
                    + (workbookDescriptor ?? string.Empty)
                    + ",window="
                    + (windowDescriptor ?? string.Empty)
                    + ",activeSheet=\""
                    + (context.ActiveSheetCodeName ?? string.Empty)
                    + "\"";
            }

            internal static string FormatHostDescriptor(TaskPaneHost host, string windowDescriptor)
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
                    + (windowDescriptor ?? string.Empty);
            }

            internal static string FormatWindowDescriptor(Excel.Window window)
            {
                return "hwnd=\""
                    + SafeGetWindowKey(window)
                    + "\",caption=\""
                    + SafeWindowCaption(window)
                    + "\"";
            }

            private static string GetPaneRoleName(TaskPaneHost host)
            {
                if (host == null || host.Control == null)
                {
                    return "Unknown";
                }

                if (IsCaseHost(host))
                {
                    return "Case";
                }

                if (IsKernelHost(host))
                {
                    return "Kernel";
                }

                if (host.Control is AccountingNavigationControl)
                {
                    return "Accounting";
                }

                return host.Control.GetType().Name;
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
        }

        internal sealed class TaskPaneManagerTestHooks
        {
            internal Action<string, string> OnHideHost { get; set; }

            internal Func<string, string, bool> TryShowHost { get; set; }

            internal Action<string> OnCasePaneUpdatedNotification { get; set; }
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
