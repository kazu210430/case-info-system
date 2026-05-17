using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Tools;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;
using RibbonExtensibility = Microsoft.Office.Core.IRibbonExtensibility;

namespace CaseInfoSystem.ExcelAddIn
{
    [ComVisible(true)]
    // VSTO / COM / UI イベントの境界クラス。業務判断は既存 service / coordinator へ委譲する。
    public partial class ThisAddIn
    {
        // 定数
        private const string TaskPaneTitle = "案件情報System";
        private static readonly bool DisableSheetActivateForFreezeIsolation = false;
        private static readonly bool DisableSheetSelectionChangeForFreezeIsolation = false;
        private static readonly bool DisableSheetChangeForFreezeIsolation = false;
        private static readonly bool DisableCaseWordWarmupForFreezeIsolation = true;
        private const string KernelSheetCommandSheetCodeName = "shCaseList";
        private const string KernelSheetCommandCellAddress = "AT1";
        private const string ProductTitle = "案件情報System";
        private const int WordWarmupDelayMs = 1500;
        private VstoEventAdapter _vstoEventAdapter;
        private HomeTransitionAdapter _homeTransitionAdapter;
        private TaskPaneEntryAdapter _taskPaneEntryAdapter;
        private AutomationSurfaceAdapter _automationSurfaceAdapter;
        private ShutdownCleanupAdapter _shutdownCleanupAdapter;
        // 基盤
        private Logger _logger;
        private ExcelInteropService _excelInteropService;
        private WorkbookRoleResolver _workbookRoleResolver;
        private CaseWorkbookOpenStrategy _caseWorkbookOpenStrategy;
        private ManagedWorkbookCloseMarkerStore _managedWorkbookCloseMarkerStore;
        private AddInStartupBoundaryCoordinator _startupBoundaryCoordinator;
        private AddInExecutionBoundaryCoordinator _executionBoundaryCoordinator;
        private AddInRuntimeExecutionDiagnosticsService _runtimeExecutionDiagnosticsService;

        // 文書実行
        private DocumentExecutionModeService _documentExecutionModeService;
        private WordInteropService _wordInteropService;

        // workbook ライフサイクル
        private KernelWorkbookService _kernelWorkbookService;
        private KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private SheetEventCoordinator _sheetEventCoordinator;
        private WorkbookLifecycleCoordinator _workbookLifecycleCoordinator;

        // Kernel 操作
        private KernelCaseCreationCommandService _kernelCaseCreationCommandService;
        private KernelCommandService _kernelCommandService;
        private KernelUserDataReflectionService _kernelUserDataReflectionService;
        private WorkbookRibbonCommandService _workbookRibbonCommandService;
        private WorkbookCaseTaskPaneRefreshCommandService _workbookCaseTaskPaneRefreshCommandService;
        private WorkbookResetCommandService _workbookResetCommandService;
        // UI / Pane 調停
        private WorkbookEventCoordinator _workbookEventCoordinator;
        private KernelWorkbookAvailabilityService _kernelWorkbookAvailabilityService;
        private KernelHomeCasePaneSuppressionCoordinator _kernelHomeCasePaneSuppressionCoordinator;
        private ExternalWorkbookDetectionService _externalWorkbookDetectionService;
        private WindowActivatePaneHandlingService _windowActivatePaneHandlingService;
        private TaskPaneRefreshOrchestrationService _taskPaneRefreshOrchestrationService;
        private TaskPaneManager _taskPaneManager;
        private KernelHomeFormHost _kernelHomeFormHost;
        private KernelCaseInteractionState _kernelCaseInteractionState;

        // COM / warm-up
        private KernelAutomationService _kernelAutomationService;
        private Timer _wordWarmupTimer;
        private bool _wordWarmupScheduled;
        private bool _wordWarmupCompleted;

        // 内部公開: 境界クラスから他コンポーネントへ渡す最小限の状態 / 判定
        internal KernelCaseInteractionState KernelCaseInteractionState
        {
            get { return _kernelCaseInteractionState; }
        }

        internal Logger Logger
        {
            get { return _logger; }
        }

        internal static string GetPrimaryTraceLogRelativePath()
        {
            return ExcelAddInTraceLogWriter.GetPrimaryTraceLogRelativePath();
        }

        internal static string GetPrimaryTraceLogPath()
        {
            return ExcelAddInTraceLogWriter.GetPrimaryTraceLogPath();
        }

        internal bool ShouldShowKernelHomeOnStartup(Excel.Workbook workbook)
        {
            return _kernelWorkbookService.ShouldShowHomeOnStartup(workbook);
        }

        internal string GetWorkbookFullNameForLogging(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return _kernelWorkbookService.IsKernelWorkbook(workbook);
        }

        // VSTO ライフサイクル
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            InitializeStartupDiagnostics();

            var compositionRoot = CreateStartupCompositionRoot();
            compositionRoot.Compose();
            ApplyCompositionRoot(compositionRoot);
            InitializeStartupBoundaryCoordinator();
            InitializeAdaptersAfterComposition();

            // 順序維持: event hook の後に startup 時の HOME 表示判定と初回 pane refresh を行う。
            _logger.Info("ThisAddIn_Startup fired.");
            HookApplicationEvents();
            _startupBoundaryCoordinator.RunAfterApplicationEventsHooked();
        }

        private void InitializeStartupDiagnostics()
        {
            _logger = new Logger(ExcelAddInTraceLogWriter.Write);
            ExcelProcessLaunchContextTracer.Trace(_logger);
            AddInDeploymentDiagnosticsTracer.Trace(_logger);
            _runtimeExecutionDiagnosticsService = new AddInRuntimeExecutionDiagnosticsService(_logger);
            _vstoEventAdapter = new VstoEventAdapter(
                Application,
                _logger,
                !DisableSheetActivateForFreezeIsolation,
                !DisableSheetSelectionChangeForFreezeIsolation,
                !DisableSheetChangeForFreezeIsolation);
            _taskPaneEntryAdapter = new TaskPaneEntryAdapter(
                _logger,
                _vstoEventAdapter.FormatActiveExcelState,
                (window, control) => CustomTaskPanes.Add(control, TaskPaneTitle, window),
                pane => CustomTaskPanes.Remove(pane));
            _executionBoundaryCoordinator = new AddInExecutionBoundaryCoordinator(Application, _logger, _vstoEventAdapter.FormatActiveExcelState);
            _homeTransitionAdapter = new HomeTransitionAdapter(
                _logger,
                _runtimeExecutionDiagnosticsService,
                _executionBoundaryCoordinator,
                _taskPaneEntryAdapter,
                _vstoEventAdapter.FormatActiveExcelState);
            _automationSurfaceAdapter = new AutomationSurfaceAdapter(Application, _logger, ProductTitle);
        }

        private AddInCompositionRoot CreateStartupCompositionRoot()
        {
            // Composition Root から VSTO 境界で使う依存と delegate を受け取る。
            return new AddInCompositionRoot(
                this,
                Application,
                _logger,
                _executionBoundaryCoordinator,
                _executionBoundaryCoordinator,
                // UI / pane
                _taskPaneEntryAdapter.ResolveWorkbookPaneWindow,
                _taskPaneEntryAdapter.TryRefreshTaskPane,
                _taskPaneEntryAdapter.IsTaskPaneRefreshSucceeded,
                GetKernelHomeForm,
                GetTaskPaneRefreshSuppressionCount,
                // Kernel HOME / sheet command
                ShowKernelHomeFromKernelCommand,
                _vstoEventAdapter.ClearKernelSheetCommandCell,
                ComObjectReleaseService.FinalRelease,
                ShowKernelHomePlaceholderWithExternalWorkbookSuppression,
                HandleExternalWorkbookDetected,
                ShouldSuppressCasePaneRefresh,
                _taskPaneEntryAdapter.RefreshTaskPane,
                // 非同期 UI 更新
                _taskPaneEntryAdapter.RequestTaskPaneDisplayForTargetWindow,
                ScheduleWordWarmup,
                KernelSheetCommandSheetCodeName,
                KernelSheetCommandCellAddress);
        }

        private int GetTaskPaneRefreshSuppressionCount()
        {
            return _executionBoundaryCoordinator == null
                ? 0
                : _executionBoundaryCoordinator.TaskPaneRefreshSuppressionCount;
        }

        private void ApplyCompositionRoot(AddInCompositionRoot compositionRoot)
        {
            // 基盤
            _excelInteropService = compositionRoot.ExcelInteropService;
            _workbookRoleResolver = compositionRoot.WorkbookRoleResolver;
            _caseWorkbookOpenStrategy = compositionRoot.CaseWorkbookOpenStrategy;
            _managedWorkbookCloseMarkerStore = compositionRoot.ManagedWorkbookCloseMarkerStore;

            // 文書実行
            _documentExecutionModeService = compositionRoot.DocumentExecutionModeService;
            _wordInteropService = compositionRoot.WordInteropService;

            // workbook ライフサイクル
            _kernelWorkbookService = compositionRoot.KernelWorkbookService;
            _kernelWorkbookLifecycleService = compositionRoot.KernelWorkbookLifecycleService;
            _sheetEventCoordinator = compositionRoot.SheetEventCoordinator;
            _workbookLifecycleCoordinator = compositionRoot.WorkbookLifecycleCoordinator;

            // Kernel 操作
            _kernelCaseCreationCommandService = compositionRoot.KernelCaseCreationCommandService;
            _kernelCommandService = compositionRoot.KernelCommandService;
            _kernelUserDataReflectionService = compositionRoot.KernelUserDataReflectionService;
            _workbookRibbonCommandService = compositionRoot.WorkbookRibbonCommandService;
            _workbookCaseTaskPaneRefreshCommandService = compositionRoot.WorkbookCaseTaskPaneRefreshCommandService;
            _workbookResetCommandService = compositionRoot.WorkbookResetCommandService;
            _kernelHomeFormHost = new KernelHomeFormHost(_kernelWorkbookService, _kernelCaseCreationCommandService, _logger);
            // UI / Pane 調停
            _workbookEventCoordinator = compositionRoot.WorkbookEventCoordinator;
            _kernelWorkbookAvailabilityService = compositionRoot.KernelWorkbookAvailabilityService;
            _kernelHomeCasePaneSuppressionCoordinator = compositionRoot.KernelHomeCasePaneSuppressionCoordinator;
            _externalWorkbookDetectionService = compositionRoot.ExternalWorkbookDetectionService;
            _windowActivatePaneHandlingService = compositionRoot.WindowActivatePaneHandlingService;
            _taskPaneRefreshOrchestrationService = compositionRoot.TaskPaneRefreshOrchestrationService;
            _taskPaneManager = compositionRoot.TaskPaneManager;
            _kernelCaseInteractionState = compositionRoot.KernelCaseInteractionState;
        }

        private void InitializeAdaptersAfterComposition()
        {
            _taskPaneEntryAdapter.Configure(
                _excelInteropService,
                _workbookRoleResolver,
                _taskPaneManager,
                _taskPaneRefreshOrchestrationService);
            _homeTransitionAdapter.Configure(
                _kernelHomeFormHost,
                _kernelWorkbookService,
                _kernelWorkbookAvailabilityService,
                _kernelHomeCasePaneSuppressionCoordinator,
                _externalWorkbookDetectionService,
                _taskPaneManager,
                _excelInteropService);
            _automationSurfaceAdapter.Configure(
                _homeTransitionAdapter,
                _excelInteropService,
                _workbookRoleResolver,
                _kernelWorkbookService,
                _kernelCommandService,
                _kernelUserDataReflectionService,
                _workbookRibbonCommandService,
                _workbookCaseTaskPaneRefreshCommandService,
                _workbookResetCommandService);
            _vstoEventAdapter.Configure(
                _startupBoundaryCoordinator,
                _excelInteropService,
                _kernelWorkbookService,
                _kernelWorkbookLifecycleService,
                _workbookLifecycleCoordinator,
                _workbookEventCoordinator,
                _sheetEventCoordinator,
                _windowActivatePaneHandlingService);
            _shutdownCleanupAdapter = new ShutdownCleanupAdapter(
                _logger,
                Application,
                UnhookApplicationEvents,
                _taskPaneEntryAdapter.StopPendingPaneRefreshTimer,
                _homeTransitionAdapter.CloseOnShutdown,
                _taskPaneEntryAdapter.DisposeAll,
                StopWordWarmupTimer,
                () => _startupBoundaryCoordinator?.StopManagedCloseStartupGuardTimer(),
                () => _caseWorkbookOpenStrategy?.ShutdownHiddenApplicationCache(),
                () => CustomTaskPanes != null,
                () => CustomTaskPanes.Count);
        }

        private void InitializeStartupBoundaryCoordinator()
        {
            _startupBoundaryCoordinator = new AddInStartupBoundaryCoordinator(
                Application,
                _logger,
                _managedWorkbookCloseMarkerStore,
                () => _kernelWorkbookService != null && _kernelWorkbookService.ShouldShowHomeOnStartup(),
                () => _kernelWorkbookService == null ? string.Empty : _kernelWorkbookService.DescribeStartupState(),
                reason => _kernelWorkbookService?.ClearHomeWorkbookBinding(reason),
                () => ShowKernelHomePlaceholder(),
                RefreshTaskPane,
                () => _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook(),
                workbook => _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookName(workbook),
                workbook => _kernelWorkbookService != null && _kernelWorkbookService.IsKernelWorkbook(workbook));
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_shutdownCleanupAdapter != null)
            {
                _shutdownCleanupAdapter.HandleShutdown();
                return;
            }

            ShutdownCleanupAdapter.WriteFallbackShutdownTrace();
        }

        private void TraceGeneratedOnShutdownBoundary(string phase)
        {
            if (_shutdownCleanupAdapter != null)
            {
                _shutdownCleanupAdapter.TraceGeneratedOnShutdownBoundary(phase);
                return;
            }

            ShutdownCleanupAdapter.TraceGeneratedOnShutdownBoundaryFallback(
                phase,
                () => CustomTaskPanes != null,
                () => CustomTaskPanes.Count);
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        // Ribbon 作成は VSTO 境界の責務としてここで引き受ける。
        protected override RibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[]
            {
                new CaseInfoSystemRibbon()
            });
        }

        // Excel application event の購読順は既存挙動維持のため変更しない。
        private void HookApplicationEvents()
        {
            _vstoEventAdapter?.SubscribeApplicationEvents();
        }

        private void UnhookApplicationEvents()
        {
            _vstoEventAdapter?.UnsubscribeApplicationEvents();
        }

        private KernelHomeForm GetKernelHomeForm()
        {
            return _homeTransitionAdapter == null ? null : _homeTransitionAdapter.GetKernelHomeForm();
        }

        internal void HandleWindowActivateEvent(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            _vstoEventAdapter?.HandleWindowActivateEvent(triggerFacts);
        }

        internal void HandleWindowActivateEvent(Excel.Workbook workbook, Excel.Window window)
        {
            _vstoEventAdapter?.HandleWindowActivateEvent(workbook, window);
        }

        internal void RequestTaskPaneDisplayForTargetWindow(TaskPaneDisplayRequest request, Excel.Workbook workbook, Excel.Window targetWindow)
        {
            _taskPaneEntryAdapter.RequestTaskPaneDisplayForTargetWindow(request, workbook, targetWindow);
        }

        private void RefreshTaskPane(TaskPaneDisplayRequest request, Excel.Workbook workbook, Excel.Window window)
        {
            _taskPaneEntryAdapter.RefreshTaskPane(request, workbook, window);
        }

        private void RefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            _taskPaneEntryAdapter.RefreshTaskPane(reason, workbook, window);
        }

        private TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneEntryAdapter.TryRefreshTaskPane(reason, workbook, window);
        }

        private bool IsTaskPaneRefreshSucceeded(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneEntryAdapter.IsTaskPaneRefreshSucceeded(reason, workbook, window);
        }

        internal void RefreshActiveTaskPane(string reason)
        {
            _taskPaneEntryAdapter.RefreshActiveTaskPane(reason);
        }

        internal void ScheduleActiveTaskPaneRefresh(string reason)
        {
            _taskPaneEntryAdapter.ScheduleActiveTaskPaneRefresh(reason);
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
            _taskPaneEntryAdapter.ScheduleWorkbookTaskPaneRefresh(workbook, reason);
        }

        internal void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            _taskPaneEntryAdapter.ShowWorkbookTaskPaneWhenReady(workbook, reason);
        }

        internal CustomTaskPane CreateTaskPane(Excel.Window window, System.Windows.Forms.UserControl control)
        {
            return _taskPaneEntryAdapter.CreateTaskPane(window, control);
        }

        internal void RemoveTaskPane(CustomTaskPane pane)
        {
            _taskPaneEntryAdapter.RemoveTaskPane(pane);
        }

        private void ShowKernelHomePlaceholder(bool clearBindingOnNewSession = false)
        {
            _homeTransitionAdapter.ShowKernelHomePlaceholder(clearBindingOnNewSession);
        }

        internal bool ShowKernelSheetAndRefreshPaneFromHome(WorkbookContext context, string sheetCodeName, string reason, out Excel.Workbook displayedWorkbook)
        {
            return _homeTransitionAdapter.ShowKernelSheetAndRefreshPaneFromHome(context, sheetCodeName, reason, out displayedWorkbook);
        }

        internal void ShowKernelHomeFromAutomation()
        {
            _automationSurfaceAdapter.ShowKernelHomeFromAutomation();
        }

        private void ShowKernelHomeFromKernelCommand()
        {
            _homeTransitionAdapter.ShowKernelHomeFromKernelCommand();
        }

        internal void LogAutomationFailure(string message, Exception ex)
        {
            _automationSurfaceAdapter.LogAutomationFailure(message, ex);
        }

        public void ShowKernelHomeFromSheet()
        {
            ShowKernelHomeFromAutomation();
        }

        private void ShowKernelHomePlaceholderWithExternalWorkbookSuppression(string reason)
        {
            _homeTransitionAdapter.ShowKernelHomePlaceholderWithExternalWorkbookSuppression(reason);
        }

        public void ReflectKernelUserDataToAccountingSet()
        {
            _automationSurfaceAdapter.ReflectKernelUserDataToAccountingSet();
        }

        public void ReflectKernelUserDataToBaseHome()
        {
            _automationSurfaceAdapter.ReflectKernelUserDataToBaseHome();
        }

        // Ribbon / COM 公開入口
        public void ShowActiveWorkbookCustomDocumentProperties()
        {
            _automationSurfaceAdapter.ShowActiveWorkbookCustomDocumentProperties();
        }

        public void SelectAndSaveActiveWorkbookSystemRoot()
        {
            _automationSurfaceAdapter.SelectAndSaveActiveWorkbookSystemRoot();
        }

        public void RefreshActiveWorkbookCaseTaskPane()
        {
            _automationSurfaceAdapter.RefreshActiveWorkbookCaseTaskPane();
        }

        public void CopySampleColumnBToHome()
        {
            _automationSurfaceAdapter.CopySampleColumnBToHome();
        }

        public void UpdateBaseDefinitionFromRibbon()
        {
            _automationSurfaceAdapter.UpdateBaseDefinitionFromRibbon();
        }

        public void ResetActiveWorkbookForDistribution()
        {
            _automationSurfaceAdapter.ResetActiveWorkbookForDistribution();
        }

        protected override object RequestComAddInAutomationService()
        {
            if (_kernelAutomationService == null)
            {
                _kernelAutomationService = new KernelAutomationService(this);
            }

            _logger?.Info("COM automation service requested.");
            if (_logger == null)
            {
                ExcelAddInTraceLogWriter.Write("COM automation service requested before startup.");
            }
            return _kernelAutomationService;
        }

        internal void HandleKernelWorkbookBecameAvailable(string eventName, Excel.Workbook workbook)
        {
            _homeTransitionAdapter.HandleKernelWorkbookBecameAvailable(eventName, workbook);
        }

        private void HandleExternalWorkbookDetected(Excel.Workbook workbook, string eventName)
        {
            _homeTransitionAdapter.HandleExternalWorkbookDetected(workbook, eventName);
        }

        internal void SuppressUpcomingKernelHomeDisplay(string reason, bool suppressOnOpen, bool suppressOnActivate)
        {
            _homeTransitionAdapter.SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen, suppressOnActivate);
        }

        internal bool ShouldSuppressKernelHomeDisplay(string eventName)
        {
            return _homeTransitionAdapter.ShouldSuppressKernelHomeDisplay(eventName);
        }

        internal void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason)
        {
            _homeTransitionAdapter.SuppressUpcomingCasePaneActivationRefresh(workbookFullName, reason);
        }

        internal void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            _homeTransitionAdapter.BeginCaseWorkbookActivateProtection(workbook, window, reason);
        }

        internal bool ShouldIgnoreWorkbookActivateDuringCaseProtection(Excel.Workbook workbook)
        {
            return _homeTransitionAdapter.ShouldIgnoreWorkbookActivateDuringCaseProtection(workbook);
        }

        internal bool ShouldIgnoreWindowActivateDuringCaseProtection(Excel.Workbook workbook, Excel.Window window)
        {
            return _homeTransitionAdapter.ShouldIgnoreWindowActivateDuringCaseProtection(workbook, window);
        }

        internal bool ShouldIgnoreTaskPaneRefreshDuringCaseProtection(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _homeTransitionAdapter.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneEntryAdapter.HasVisibleCasePaneForWorkbookWindow(workbook, window);
        }

        private bool ShouldSuppressCasePaneRefresh(string eventName, Excel.Workbook workbook)
        {
            return _homeTransitionAdapter.ShouldSuppressCasePaneRefresh(eventName, workbook);
        }

        internal bool IsKernelHomeSuppressionActive(string eventName, bool consume)
        {
            return _homeTransitionAdapter.IsKernelHomeSuppressionActive(eventName, consume);
        }

        // Word warm-up
        private void ScheduleWordWarmup()
        {
            if (DisableCaseWordWarmupForFreezeIsolation)
            {
                _logger?.Info("Word warm-up skipped for freeze isolation.");
                return;
            }

            if (_wordInteropService == null)
            {
                return;
            }

            if (_wordWarmupCompleted)
            {
                return;
            }

            if (_documentExecutionModeService == null || !_documentExecutionModeService.IsWordWarmupEnabled())
            {
                return;
            }

            if (_wordWarmupTimer == null)
            {
                _wordWarmupTimer = new Timer();
                _wordWarmupTimer.Interval = WordWarmupDelayMs;
                _wordWarmupTimer.Tick += WordWarmupTimer_Tick;
            }

            if (_wordWarmupScheduled)
            {
                return;
            }

            _wordWarmupScheduled = true;
            _wordWarmupTimer.Interval = WordWarmupDelayMs;
            _wordWarmupTimer.Stop();
            _wordWarmupTimer.Start();
        }

        private void StopWordWarmupTimer()
        {
            if (_wordWarmupTimer == null)
            {
                return;
            }

            _wordWarmupTimer.Stop();
        }

        private void WordWarmupTimer_Tick(object sender, EventArgs e)
        {
            StopWordWarmupTimer();
            _wordWarmupScheduled = false;

            if (_wordWarmupCompleted || _wordInteropService == null)
            {
                return;
            }

            try
            {
                _wordInteropService.WarmUpApplication();
                _wordWarmupCompleted = true;
            }
            catch (Exception ex)
            {
                _logger.Error("Word warm-up failed.", ex);
            }
        }
    }
}





