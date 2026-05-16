using System;
using System.Diagnostics;
using System.Globalization;
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
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private static readonly bool DisableSheetActivateForFreezeIsolation = false;
        private static readonly bool DisableSheetSelectionChangeForFreezeIsolation = false;
        private static readonly bool DisableSheetChangeForFreezeIsolation = false;
        private static readonly bool DisableCaseWordWarmupForFreezeIsolation = true;
        private const string KernelSheetCommandSheetCodeName = "shCaseList";
        private const string KernelSheetCommandCellAddress = "AT1";
        private const string ProductTitle = "案件情報System";
        private const int WordWarmupDelayMs = 1500;
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
        private ApplicationEventSubscriptionService _applicationEventSubscriptionService;
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
        private KernelHomeCoordinator _kernelHomeCoordinator;
        private KernelHomeCasePaneSuppressionCoordinator _kernelHomeCasePaneSuppressionCoordinator;
        private ExternalWorkbookDetectionService _externalWorkbookDetectionService;
        private WindowActivatePaneHandlingService _windowActivatePaneHandlingService;
        private TaskPaneRefreshOrchestrationService _taskPaneRefreshOrchestrationService;
        private TaskPaneManager _taskPaneManager;
        private KernelHomeFormHost _kernelHomeFormHost;
        private KernelCaseInteractionState _kernelCaseInteractionState;
        private int _kernelFlickerTraceRefreshCallSequence;

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
            InitializeApplicationEventSubscriptionService();
            InitializeStartupBoundaryCoordinator();

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
            _executionBoundaryCoordinator = new AddInExecutionBoundaryCoordinator(Application, _logger, FormatActiveExcelState);
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
                ResolveWorkbookPaneWindow,
                TryRefreshTaskPane,
                IsTaskPaneRefreshSucceeded,
                GetKernelHomeForm,
                GetTaskPaneRefreshSuppressionCount,
                // Kernel HOME / sheet command
                ShowKernelHomeFromKernelCommand,
                ClearKernelSheetCommandCell,
                ComObjectReleaseService.FinalRelease,
                ShowKernelHomePlaceholderWithExternalWorkbookSuppression,
                HandleExternalWorkbookDetected,
                ShouldSuppressCasePaneRefresh,
                RefreshTaskPane,
                // 非同期 UI 更新
                RequestTaskPaneDisplayForTargetWindow,
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
            _kernelHomeCoordinator = compositionRoot.KernelHomeCoordinator;
            _kernelHomeCasePaneSuppressionCoordinator = compositionRoot.KernelHomeCasePaneSuppressionCoordinator;
            _externalWorkbookDetectionService = compositionRoot.ExternalWorkbookDetectionService;
            _windowActivatePaneHandlingService = compositionRoot.WindowActivatePaneHandlingService;
            _taskPaneRefreshOrchestrationService = compositionRoot.TaskPaneRefreshOrchestrationService;
            _taskPaneManager = compositionRoot.TaskPaneManager;
            _kernelCaseInteractionState = compositionRoot.KernelCaseInteractionState;
        }

        private void InitializeApplicationEventSubscriptionService()
        {
            _applicationEventSubscriptionService = new ApplicationEventSubscriptionService(
                Application,
                Application_WorkbookOpen,
                Application_WorkbookActivate,
                Application_WorkbookBeforeSave,
                Application_WorkbookBeforeClose,
                Application_WindowActivate,
                Application_SheetActivate,
                Application_SheetSelectionChange,
                Application_SheetChange,
                Application_AfterCalculate,
                !DisableSheetActivateForFreezeIsolation,
                !DisableSheetSelectionChangeForFreezeIsolation,
                !DisableSheetChangeForFreezeIsolation);
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
			if (_logger != null)
			{
				_logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-entry");
				LogShutdownState("handler-entry");
				RunShutdownStep("UnhookApplicationEvents", UnhookApplicationEvents);
				RunShutdownStep(
					"StopPendingPaneRefreshTimer",
					() => _taskPaneRefreshOrchestrationService?.StopPendingPaneRefreshTimer());
				RunShutdownStep("KernelHomeFormHost.CloseOnShutdown", () => _kernelHomeFormHost?.CloseOnShutdown());
				RunShutdownStep(
					"TaskPaneManager.DisposeAll",
					() =>
					{
						LogShutdownState("before-taskpane-manager-disposeall");
						_taskPaneManager?.DisposeAll();
						LogShutdownState("after-taskpane-manager-disposeall");
					});
				RunShutdownStep("StopWordWarmupTimer", StopWordWarmupTimer);
				RunShutdownStep("StopManagedCloseStartupGuardTimer", () => _startupBoundaryCoordinator?.StopManagedCloseStartupGuardTimer());
				RunShutdownStep(
					"ShutdownHiddenApplicationCache",
					() => _caseWorkbookOpenStrategy?.ShutdownHiddenApplicationCache());

				LogShutdownState("before-generated-base-boundary");
				_logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-before-generated-base-boundary");
				_logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-exit");
				_logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-complete");
				_logger.Info("ThisAddIn_Shutdown fired.");
				return;
			}

			ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-entry logger=null");
			ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-before-generated-base-boundary logger=null");
			ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-exit logger=null");
			ExcelAddInTraceLogWriter.Write("[KernelFlickerTrace] source=ThisAddIn action=shutdown-handler-complete logger=null");
			ExcelAddInTraceLogWriter.Write("ThisAddIn_Shutdown fired.");
		}

		private void RunShutdownStep(string stepName, Action action)
		{
			_logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-step-start step=" + stepName);

			try
			{
				action?.Invoke();
				_logger.Info("[KernelFlickerTrace] source=ThisAddIn action=shutdown-step-complete step=" + stepName);
			}
			catch (Exception exception)
			{
				_logger.Error(
					"[KernelFlickerTrace] source=ThisAddIn action=shutdown-step-failure step="
						+ stepName
						+ ", exceptionType="
						+ exception.GetType().Name
						+ ", hResult=0x"
						+ exception.HResult.ToString("X8", CultureInfo.InvariantCulture)
						+ ", message="
						+ exception.Message,
					exception);
				throw;
			}
		}

		private void LogShutdownState(string phase)
		{
			SafeWriteShutdownTrace(
				KernelFlickerTracePrefix
				+ " source=ThisAddIn action=shutdown-state phase="
				+ (phase ?? string.Empty)
				+ ", "
				+ CaptureShutdownExcelStateFacts()
				+ ", "
				+ CaptureCustomTaskPaneFacts());
		}

		private string CaptureShutdownExcelStateFacts()
		{
			string applicationVisible = ReadShutdownValue(() => Application.Visible, out bool applicationVisibleReadFailed);
			string workbooksCount = ReadShutdownValue(() => Application.Workbooks.Count, out bool workbooksCountReadFailed);
			string windowsCount = ReadShutdownValue(() => Application.Windows.Count, out bool windowsCountReadFailed);
			string displayAlerts = ReadShutdownValue(() => Application.DisplayAlerts, out bool displayAlertsReadFailed);
			string enableEvents = ReadShutdownValue(() => Application.EnableEvents, out bool enableEventsReadFailed);
			string calculationState = ReadShutdownValue(() => Application.CalculationState, out bool calculationStateReadFailed);
			string hwnd = ReadShutdownValue(() => Application.Hwnd, out bool hwndReadFailed);

			return "pid="
				+ Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
				+ ", applicationPresent="
				+ (Application != null).ToString()
				+ ", applicationVisible="
				+ applicationVisible
				+ ", applicationVisibleReadFailed="
				+ applicationVisibleReadFailed.ToString()
				+ ", workbooksCount="
				+ workbooksCount
				+ ", workbooksCountReadFailed="
				+ workbooksCountReadFailed.ToString()
				+ ", "
				+ CaptureShutdownActiveWorkbookFacts()
				+ ", windowsCount="
				+ windowsCount
				+ ", windowsCountReadFailed="
				+ windowsCountReadFailed.ToString()
				+ ", displayAlerts="
				+ displayAlerts
				+ ", displayAlertsReadFailed="
				+ displayAlertsReadFailed.ToString()
				+ ", enableEvents="
				+ enableEvents
				+ ", enableEventsReadFailed="
				+ enableEventsReadFailed.ToString()
				+ ", calculationState="
				+ calculationState
				+ ", calculationStateReadFailed="
				+ calculationStateReadFailed.ToString()
				+ ", hwnd="
				+ hwnd
				+ ", hwndReadFailed="
				+ hwndReadFailed.ToString();
		}

		private string CaptureShutdownActiveWorkbookFacts()
		{
			bool activeWorkbookPresent = false;
			bool activeWorkbookReadFailed = false;
			bool activeWorkbookNameReadFailed = false;
			string activeWorkbookName = string.Empty;

			try
			{
				Excel.Workbook activeWorkbook = Application == null ? null : Application.ActiveWorkbook;
				activeWorkbookPresent = activeWorkbook != null;
				if (activeWorkbook != null)
				{
					try
					{
						activeWorkbookName = SanitizeShutdownLogValue(activeWorkbook.Name);
					}
					catch (Exception)
					{
						activeWorkbookNameReadFailed = true;
					}
				}
			}
			catch (Exception)
			{
				activeWorkbookReadFailed = true;
			}

			return "activeWorkbookPresent="
				+ activeWorkbookPresent.ToString()
				+ ", activeWorkbookName=\""
				+ activeWorkbookName
				+ "\", activeWorkbookReadFailed="
				+ activeWorkbookReadFailed.ToString()
				+ ", activeWorkbookNameReadFailed="
				+ activeWorkbookNameReadFailed.ToString();
		}

		private string CaptureCustomTaskPaneFacts()
		{
			string customTaskPanesCount = ReadShutdownValue(() => CustomTaskPanes.Count, out bool customTaskPanesCountReadFailed);
			return "customTaskPanesPresent="
				+ (CustomTaskPanes != null).ToString()
				+ ", customTaskPanesCount="
				+ customTaskPanesCount
				+ ", customTaskPanesCountReadFailed="
				+ customTaskPanesCountReadFailed.ToString();
		}

		private static string ReadShutdownValue<T>(Func<T> read, out bool readFailed)
		{
			readFailed = false;
			try
			{
				T value = read();
				return SanitizeShutdownLogValue(value == null ? string.Empty : value.ToString());
			}
			catch (Exception)
			{
				readFailed = true;
				return string.Empty;
			}
		}

		private static string SanitizeShutdownLogValue(string value)
		{
			return (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
		}

		private void TraceGeneratedOnShutdownBoundary(string phase)
		{
			SafeWriteShutdownTrace(
				KernelFlickerTracePrefix
				+ " source=ThisAddIn action=generated-onshutdown-boundary phase="
				+ (phase ?? string.Empty)
				+ ", pid="
				+ Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
				+ ", "
				+ CaptureCustomTaskPaneFacts());
		}

		private void SafeWriteShutdownTrace(string message)
		{
			try
			{
				if (_logger != null)
				{
					_logger.Info(message);
					return;
				}

				ExcelAddInTraceLogWriter.Write((message ?? string.Empty) + " logger=null");
			}
			catch (Exception)
			{
				// Shutdown diagnostics must never change the unload control flow.
			}
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
            _applicationEventSubscriptionService?.Subscribe();
        }

        private void UnhookApplicationEvents()
        {
            _applicationEventSubscriptionService?.Unsubscribe();
        }

        private KernelHomeForm GetKernelHomeForm()
        {
            return _kernelHomeFormHost == null ? null : _kernelHomeFormHost.Current;
        }

        // Excel application event handler
        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            _startupBoundaryCoordinator?.MarkWorkbookOpenObserved();
            EnsureKernelFlickerTraceForWorkbookOpen(workbook);
            EventBoundaryGuard.Execute(_logger, nameof(Application_WorkbookOpen), () => _workbookLifecycleCoordinator?.OnWorkbookOpen(workbook));
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_WorkbookActivate), () => _workbookLifecycleCoordinator?.OnWorkbookActivate(workbook));
        }

        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_WindowActivate), () =>
            {
                WindowActivateTaskPaneTriggerFacts triggerFacts = CaptureWindowActivateTaskPaneTriggerFacts(
                    workbook,
                    window,
                    "ThisAddIn.Application_WindowActivate");
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=ExcelEventBoundary action=fire event=WindowActivate workbook="
                    + triggerFacts.WorkbookDescriptor
                    + ", eventWindow="
                    + triggerFacts.WindowDescriptor
                    + ", activeState="
                    + triggerFacts.ActiveState
                    + ", triggerRole=TaskPaneDisplayRefreshTrigger");
                _logger?.Info(
                    "Excel WindowActivate fired. workbook="
                    + triggerFacts.WorkbookFullName
                    + ", windowHwnd="
                    + triggerFacts.WindowHwnd
                    + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
                _workbookEventCoordinator.OnWindowActivate(triggerFacts);
            });
        }

        internal void HandleWindowActivateEvent(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            if (triggerFacts == null)
            {
                triggerFacts = CaptureWindowActivateTaskPaneTriggerFacts(
                    null,
                    null,
                    "ThisAddIn.HandleWindowActivateEvent.NullFacts");
            }

            Excel.Workbook workbook = triggerFacts.Workbook;
            Excel.Window window = triggerFacts.Window;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WindowActivate triggerRole=TaskPaneDisplayRefreshTrigger workbook="
                + triggerFacts.WorkbookDescriptor
                + ", eventWindow="
                + triggerFacts.WindowDescriptor
                + ", activeState="
                + triggerFacts.ActiveState
                + ", captureOwner="
                + triggerFacts.CaptureOwner);
            _logger?.Info("TaskPane event entry. event=WindowActivate, workbook=" + SafeWorkbookFullName(workbook) + ", windowHwnd=" + SafeWindowHwnd(window) + ", activeWorkbook=" + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook()) + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
            NewCaseVisibilityObservation.Log(_logger, _excelInteropService, Application, workbook, window, "WindowActivate-event", "ThisAddIn.HandleWindowActivateEvent");
            _windowActivatePaneHandlingService?.Handle(triggerFacts);
        }

        internal void HandleWindowActivateEvent(Excel.Workbook workbook, Excel.Window window)
        {
            HandleWindowActivateEvent(CaptureWindowActivateTaskPaneTriggerFacts(
                workbook,
                window,
                "ThisAddIn.HandleWindowActivateEvent.Legacy"));
        }

        private void Application_SheetActivate(object sh)
        {
            _logger?.Debug("Application_SheetActivate", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetActivate), () => _sheetEventCoordinator?.OnSheetActivate(sh));
            _logger?.Debug("Application_SheetActivate", "returned.");
        }

        private void Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            _logger?.Debug("Application_SheetSelectionChange", "entry.");
            if (!(sh is Excel.Worksheet) || target == null)
            {
                _logger?.Debug("Application_SheetSelectionChange", "returned.");
                return;
            }

            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetSelectionChange), () =>
            {
                _sheetEventCoordinator?.OnSheetSelectionChange(sh, target);
            });
            _logger?.Debug("Application_SheetSelectionChange", "returned.");
        }

        private void Application_SheetChange(object sh, Excel.Range target)
        {
            _logger?.Debug("Application_SheetChange", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetChange), () => _sheetEventCoordinator?.OnSheetChange(sh, target));
            _logger?.Debug("Application_SheetChange", "returned.");
        }

        private void Application_AfterCalculate()
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_AfterCalculate), () => _sheetEventCoordinator?.OnAfterCalculate(Application));
            _logger?.Debug("Application_AfterCalculate", "EventBoundaryGuard.Execute returned.");
        }

        private void ClearKernelSheetCommandCell(Excel.Range commandCell)
        {
            if (commandCell == null)
            {
                return;
            }

            bool previousEnableEvents = Application.EnableEvents;
            try
            {
                Application.EnableEvents = false;
                commandCell.Value2 = string.Empty;
            }
            finally
            {
                Application.EnableEvents = previousEnableEvents;
            }
        }

        // 例外を外へ出さない安全取得 helper 群
        private static string SafeSheetName(object sheetObject)
        {
            try
            {
                var worksheet = sheetObject as Excel.Worksheet;
                return worksheet == null ? string.Empty : worksheet.CodeName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeRangeAddress(Excel.Range range)
        {
            try
            {
                return range == null ? string.Empty : Convert.ToString(range.Address[false, false]) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
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

        // Excel workbook lifecycle event handler
        private void Application_WorkbookBeforeSave(Excel.Workbook workbook, bool saveAsUi, ref bool cancel)
        {
            EventBoundaryGuard.ExecuteCancelable(_logger, nameof(Application_WorkbookBeforeSave), ref cancel, HandleBeforeSave);

            void HandleBeforeSave(ref bool innerCancel)
            {
                _logger?.Info(
                    "Excel WorkbookBeforeSave fired. workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook))
                    + ", saveAsUi="
                    + saveAsUi.ToString()
                    + ", cancel="
                    + innerCancel.ToString());
                _kernelWorkbookLifecycleService?.HandleWorkbookBeforeSave(workbook, saveAsUi, ref innerCancel);
            }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            EventBoundaryGuard.ExecuteCancelable(_logger, nameof(Application_WorkbookBeforeClose), ref cancel, HandleBeforeClose);

            void HandleBeforeClose(ref bool innerCancel)
            {
                if (_workbookLifecycleCoordinator != null)
                {
                    _workbookLifecycleCoordinator.OnWorkbookBeforeClose(workbook, ref innerCancel);
                }
            }
        }

        // TaskPane display-entry boundary.
        // WindowActivate / post-action refresh から共通で入る create-side refresh/display の入口であり、
        // concrete VSTO CustomTaskPane create/remove 自体は下の CreateTaskPane(...) / RemoveTaskPane(...) に残す。
        internal void RequestTaskPaneDisplayForTargetWindow(TaskPaneDisplayRequest request, Excel.Workbook workbook, Excel.Window targetWindow)
        {
            if (request != null && request.RefreshIntent == TaskPaneDisplayRefreshIntent.ForceRefresh)
            {
                _taskPaneManager?.PrepareTargetWindowForForcedRefresh(targetWindow);
            }

            TaskPaneDisplayEntryDecision displayEntryDecision = PaneDisplayPolicy.Decide(
                request,
                _taskPaneManager,
                _workbookRoleResolver,
                workbook,
                targetWindow);
            LogTaskPaneDisplayEntryDecision(request, displayEntryDecision, workbook, targetWindow);
            switch (displayEntryDecision.Result)
            {
                case PaneDisplayPolicyResult.ShowExisting:
                    _taskPaneManager?.TryShowExistingPane(workbook, targetWindow, "DisplayRequest.ShowExisting");
                    return;

                case PaneDisplayPolicyResult.Hide:
                    _taskPaneManager?.HidePaneForWindow(targetWindow);
                    return;

                case PaneDisplayPolicyResult.Reject:
                    return;
            }

            RefreshTaskPane(request, workbook, targetWindow);
        }

        private void LogTaskPaneDisplayEntryDecision(
            TaskPaneDisplayRequest request,
            TaskPaneDisplayEntryDecision decision,
            Excel.Workbook workbook,
            Excel.Window targetWindow)
        {
            if (request == null || !request.IsWindowActivateTrigger)
            {
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=window-activate-display-entry-decision reason="
                + request.ToReasonString()
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", displayEntryResult="
                + (decision == null ? PaneDisplayPolicyResult.Reject.ToString() : decision.Result.ToString())
                + ", displayCompletionOutcome=False"
                + ", recoveryOwner=False"
                + ", foregroundGuaranteeOwner=False"
                + ", hiddenExcelOwner=False"
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", targetWindow="
                + FormatWindowDescriptor(targetWindow)
                + FormatDisplayEntryStateTraceFields(decision == null ? null : decision.State)
                + FormatDisplayRequestTraceFields(request));
        }

        private static string FormatDisplayEntryStateTraceFields(TaskPaneDisplayEntryState state)
        {
            if (state == null)
            {
                return ", displayEntryState=null";
            }

            return ", displayEntryState=present"
                + ", hasTargetWindow=" + state.HasTargetWindow.ToString()
                + ", hasResolvableWindowKey=" + state.HasResolvableWindowKey.ToString()
                + ", hasManagedPane=" + state.HasManagedPane.ToString()
                + ", hasExistingHost=" + state.HasExistingHost.ToString()
                + ", isSameWorkbook=" + state.IsSameWorkbook.ToString()
                + ", isRenderSignatureCurrent=" + state.IsRenderSignatureCurrent.ToString();
        }

        private void RefreshTaskPane(TaskPaneDisplayRequest request, Excel.Workbook workbook, Excel.Window window)
        {
            string reason = request == null ? string.Empty : request.ToReasonString();
            RefreshTaskPane(reason, workbook, window, request);
        }

        private void RefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            RefreshTaskPane(reason, workbook, window, request: null);
        }

        private void RefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window, TaskPaneDisplayRequest request)
        {
            int refreshCallId = ++_kernelFlickerTraceRefreshCallSequence;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=refresh-call-start refreshCallId="
                + refreshCallId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", activeState="
                + FormatActiveExcelState()
                + FormatDisplayRequestTraceFields(request));
            TaskPaneRefreshAttemptResult result = request == null
                ? TryRefreshTaskPane(reason, workbook, window)
                : _taskPaneRefreshOrchestrationService.TryRefreshTaskPane(request, workbook, window);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=refresh-call-end refreshCallId="
                + refreshCallId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", result="
                + (result == null ? "null" : result.IsRefreshSucceeded.ToString())
                + FormatDisplayRequestTraceFields(request));
        }

        private static string FormatDisplayRequestTraceFields(TaskPaneDisplayRequest request)
        {
            if (request == null)
            {
                return string.Empty;
            }

            string details =
                ", displayRequestSource=" + request.Source.ToString()
                + ", displayRequestRefreshIntent=" + request.RefreshIntent.ToString()
                + ", displayTriggerReason=" + request.ToReasonString();
            if (!request.IsWindowActivateTrigger)
            {
                return details;
            }

            WindowActivateTaskPaneTriggerFacts facts = request.WindowActivateTriggerFacts;
            return details
                + ", windowActivateTriggerRole=TaskPaneDisplayRefreshTrigger"
                + ", windowActivateRecoveryOwner=False"
                + ", windowActivateForegroundGuaranteeOwner=False"
                + ", windowActivateHiddenExcelOwner=False"
                + ", windowActivateCaptureOwner=" + (facts == null ? string.Empty : facts.CaptureOwner)
                + ", windowActivateWorkbookPresent=" + (facts != null && facts.HasWorkbook).ToString()
                + ", windowActivateWindowPresent=" + (facts != null && facts.HasWindow).ToString()
                + ", windowActivateWindowHwnd=" + (facts == null ? string.Empty : facts.WindowHwnd);
        }

        private TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneRefreshOrchestrationService.TryRefreshTaskPane(reason, workbook, window);
        }

        private bool IsTaskPaneRefreshSucceeded(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return TryRefreshTaskPane(reason, workbook, window).IsRefreshSucceeded;
        }

        internal void RefreshActiveTaskPane(string reason)
        {
            _taskPaneRefreshOrchestrationService.RefreshActiveTaskPane(reason);
        }

        internal void ScheduleActiveTaskPaneRefresh(string reason)
        {
            _taskPaneRefreshOrchestrationService.ScheduleActiveTaskPaneRefresh(reason);
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
            _taskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(workbook, reason);
        }

        internal void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            _taskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(workbook, reason);
        }

        // Concrete VSTO adapter boundary for CustomTaskPane creation.
        // RequestTaskPaneDisplayForTargetWindow(...) is only the display-entry; the actual VSTO create call stays here.
        // Higher-level create-side ownership stays in TaskPaneHostFactory/TaskPaneHost/TaskPaneHostRegistry/TaskPaneManager.
        internal CustomTaskPane CreateTaskPane(Excel.Window window, System.Windows.Forms.UserControl control)
        {
            CustomTaskPane pane = CustomTaskPanes.Add(control, TaskPaneTitle, window);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            return pane;
        }

        // Concrete VSTO adapter boundary for CustomTaskPane removal.
        // This remains separate from the display-entry above so create/remove adapter calls stay readable without changing timing.
        // Standard remove, replacement remove, and shutdown cleanup all reach this adapter through TaskPaneHost.Dispose().
        // Current-state remove ordering is still owned by TaskPaneHost.Dispose(), not by this method.
        internal void RemoveTaskPane(CustomTaskPane pane)
        {
            if (pane == null)
            {
                return;
            }

            CustomTaskPanes.Remove(pane);
        }

        // Kernel HOME / sheet 遷移
        private void ShowKernelHomePlaceholder(bool clearBindingOnNewSession = false)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.ResetKernelHomeExternalCloseRequested();

            _kernelHomeFormHost.GetOrCreate(clearBindingOnNewSession);
            _taskPaneManager?.HideKernelPanes();
            _kernelHomeFormHost.ReloadCurrent();

            _runtimeExecutionDiagnosticsService?.Trace("ShowKernelHomePlaceholder");
            _kernelWorkbookService.PrepareForHomeDisplayFromSheet();
            _kernelWorkbookService.EnsureHomeDisplayHidden("ThisAddIn.ShowKernelHomePlaceholder.BeforeShow");

            _kernelHomeFormHost.ShowAndActivateCurrent();
        }

        internal bool ShowKernelSheetAndRefreshPaneFromHome(WorkbookContext context, string sheetCodeName, string reason, out Excel.Workbook displayedWorkbook)
        {
            displayedWorkbook = null;
            if (context == null)
            {
                _logger?.Warn(
                    "ShowKernelSheetAndRefreshPaneFromHome skipped because workbook context was not available. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            Excel.Workbook resolvedDisplayedWorkbook = _kernelWorkbookService.ResolveKernelWorkbook(context);
            if (resolvedDisplayedWorkbook == null)
            {
                _logger?.Warn(
                    "ShowKernelSheetAndRefreshPaneFromHome skipped because bound kernel workbook could not be resolved. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            KernelFlickerTraceContext.BeginNewTrace();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=trace-begin trigger=ShowKernelSheetAndRefreshPaneFromHomeBoundContext traceOriginReason="
                + (reason ?? string.Empty)
                + ", sheetCodeName="
                + (sheetCodeName ?? string.Empty)
                + ", workbook="
                + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(resolvedDisplayedWorkbook))
                + ", activeState="
                + FormatActiveExcelState());
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            bool shouldSuspendScreenUpdating = !string.IsNullOrWhiteSpace(reason)
                && reason.IndexOf("KernelHomeForm.OpenSheet", StringComparison.OrdinalIgnoreCase) >= 0;
            bool shown = false;
            Action performTransition = () =>
            {
                HideKernelHomePlaceholder();
                shown = _kernelWorkbookService.TryShowSheetByCodeName(context, sheetCodeName, reason);
                _logger?.Info("[Transition] bound-context sheet shown=" + shown + ", sheet=" + sheetCodeName);
                if (!shown)
                {
                    return;
                }

                RefreshTaskPane(reason, resolvedDisplayedWorkbook, null);
            };

            if (shouldSuspendScreenUpdating)
            {
                _executionBoundaryCoordinator.Execute(performTransition);
            }
            else
            {
                performTransition();
            }

            if (!shown)
            {
                _logger?.Info(
                    "ShowKernelSheetAndRefreshPaneFromHome aborted because target sheet could not be shown. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            displayedWorkbook = resolvedDisplayedWorkbook;
            return true;
        }

        internal void ShowKernelHomeFromAutomation()
        {
            _logger?.Info("Kernel home requested from COM automation.");
            if (_logger == null)
            {
            ExcelAddInTraceLogWriter.Write("Kernel home requested from COM automation.");
            }
            ShowKernelHomePlaceholderWithExternalWorkbookSuppressionForNewSession("KernelAutomationService.ShowHome");
        }

        private void ShowKernelHomeFromKernelCommand()
        {
            ShowKernelHomePlaceholderWithExternalWorkbookSuppression("KernelCommandService.OpenHome");
        }

        internal void LogAutomationFailure(string message, Exception ex)
        {
            if (_logger != null)
            {
                _logger.Error(message, ex);
                return;
            }

            ExcelAddInTraceLogWriter.Write((message ?? string.Empty) + " exception=" + (ex == null ? string.Empty : ex.ToString()));
        }

        public void ShowKernelHomeFromSheet()
        {
            ShowKernelHomeFromAutomation();
        }

        // Kernel HOME の明示表示直後だけ activate 系イベントの自動クローズを抑止する。
        private void ShowKernelHomePlaceholderWithExternalWorkbookSuppression(string reason)
        {
            ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(reason, clearBindingOnNewSession: false);
        }

        private void ShowKernelHomePlaceholderWithExternalWorkbookSuppressionForNewSession(string reason)
        {
            ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(reason, clearBindingOnNewSession: true);
        }

        private void ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(string reason, bool clearBindingOnNewSession)
        {
            // 処理ブロック: HOME を明示表示した直後だけ WorkbookActivate/WindowActivate を抑止する。
            // ここでは将来の activate 系イベントに効く Kernel HOME 抑止要求を発行する。
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            ShowKernelHomePlaceholder(clearBindingOnNewSession);
        }

        public void ReflectKernelUserDataToAccountingSet()
        {
            _logger?.Info("Kernel user data reflection requested from COM automation. target=AccountingSet");
            WorkbookContext context = ResolveKernelReflectionContextForAutomation();
            _kernelUserDataReflectionService?.ReflectToAccountingSetOnly(context);
        }

        public void ReflectKernelUserDataToBaseHome()
        {
            _logger?.Info("Kernel user data reflection requested from COM automation. target=BaseHome");
            WorkbookContext context = ResolveKernelReflectionContextForAutomation();
            _kernelUserDataReflectionService?.ReflectToBaseHomeOnly(context);
        }

        // Ribbon / COM 公開入口
        public void ShowActiveWorkbookCustomDocumentProperties()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            WorkbookRibbonCommandService ribbonCommandService = _workbookRibbonCommandService;
            ribbonCommandService?.ShowCustomDocumentProperties(targetWorkbook);
        }

        public void SelectAndSaveActiveWorkbookSystemRoot()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            WorkbookRibbonCommandService ribbonCommandService = _workbookRibbonCommandService;
            ribbonCommandService?.SelectAndSaveSystemRoot(targetWorkbook);
        }

        public void RefreshActiveWorkbookCaseTaskPane()
        {
            Excel.Workbook workbook = ResolveRibbonTargetWorkbook();
            WorkbookCaseTaskPaneRefreshCommandService workbookCaseTaskPaneRefreshCommandService = _workbookCaseTaskPaneRefreshCommandService;
            if (workbookCaseTaskPaneRefreshCommandService == null)
            {
                MessageBox.Show("Pane 更新サービスを利用できません。", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            workbookCaseTaskPaneRefreshCommandService.Refresh(workbook);
        }

        public void CopySampleColumnBToHome()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            WorkbookRibbonCommandService ribbonCommandService = _workbookRibbonCommandService;
            ribbonCommandService?.CopySampleColumnBToHome(targetWorkbook);
        }

        public void UpdateBaseDefinitionFromRibbon()
        {
            KernelCommandService kernelCommandService = _kernelCommandService;
            if (kernelCommandService == null)
            {
                MessageBox.Show("Base定義更新サービスを利用できません。", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                WorkbookContext context = ResolveKernelCommandContextForRibbon();
                kernelCommandService.Execute(context, KernelNavigationActionIds.SyncBaseHomeFieldInventory);
            }
            catch (Exception ex)
            {
                _logger?.Error("UpdateBaseDefinitionFromRibbon failed.", ex);
                MessageBox.Show("Base定義更新を実行できませんでした。ログを確認してください。", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void ResetActiveWorkbookForDistribution()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            WorkbookResetCommandService workbookResetCommandService = _workbookResetCommandService;
            WorkbookResetResult result = workbookResetCommandService == null
                ? new WorkbookResetResult
                {
                    Success = false,
                    Message = "配布前リセットサービスを利用できません。"
                }
                : workbookResetCommandService.Execute(targetWorkbook);
            workbookResetCommandService?.ShowResult(result);
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

        private Excel.Workbook ResolveRibbonTargetWorkbook()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            if (activeWorkbook != null)
            {
                return activeWorkbook;
            }

            if (_excelInteropService == null)
            {
                return null;
            }

            var openWorkbooks = _excelInteropService.GetOpenWorkbooks();
            return openWorkbooks.Count == 1 ? openWorkbooks[0] : null;
        }

        private WorkbookContext ResolveKernelCommandContextForRibbon()
        {
            Excel.Workbook workbook = ResolveRibbonTargetWorkbook();
            string systemRoot = string.Empty;
            if (workbook != null && _excelInteropService != null)
            {
                systemRoot = _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT");
            }

            if ((workbook == null || string.IsNullOrWhiteSpace(systemRoot)) && _kernelWorkbookService != null)
            {
                string boundSystemRoot;
                Excel.Workbook boundWorkbook;
                if (_kernelWorkbookService.TryGetValidHomeWorkbookBinding(out boundWorkbook, out boundSystemRoot))
                {
                    workbook = boundWorkbook;
                    systemRoot = boundSystemRoot;
                }
            }

            if (workbook == null || _excelInteropService == null)
            {
                throw new InvalidOperationException("Kernel workbook context was not available for Base definition update.");
            }

            WorkbookRole role = _workbookRoleResolver == null
                ? WorkbookRole.Unknown
                : _workbookRoleResolver.Resolve(workbook);
            return new WorkbookContext(
                workbook,
                TryGetActiveWindow(),
                role,
                systemRoot,
                _excelInteropService.GetWorkbookFullName(workbook),
                _excelInteropService.GetActiveSheetCodeName(workbook));
        }

        private Excel.Window TryGetActiveWindow()
        {
            try
            {
                return Application == null ? null : Application.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private WorkbookContext ResolveKernelReflectionContextForAutomation()
        {
            Excel.Workbook workbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            string systemRoot = _excelInteropService == null || workbook == null
                ? string.Empty
                : _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT");

            if (workbook == null && _kernelWorkbookService != null)
            {
                string boundSystemRoot;
                if (_kernelWorkbookService.TryGetValidHomeWorkbookBinding(out workbook, out boundSystemRoot))
                {
                    systemRoot = boundSystemRoot;
                }
            }

            if (workbook == null || _excelInteropService == null || _workbookRoleResolver == null)
            {
                throw new InvalidOperationException("Kernel workbook context was not available for user-data reflection.");
            }

            return new WorkbookContext(
                workbook,
                _excelInteropService.GetActiveWindow(),
                _workbookRoleResolver.Resolve(workbook),
                systemRoot,
                _excelInteropService.GetWorkbookFullName(workbook),
                _excelInteropService.GetActiveSheetCodeName(workbook));
        }

        private void HideKernelHomePlaceholder()
        {
            _kernelHomeFormHost?.HideCurrent();
        }

        // Kernel workbook 到達後の UI 反映責務は ThisAddIn に残し、判定自体は coordinator へ委譲する。
        internal void HandleKernelWorkbookBecameAvailable(string eventName, Excel.Workbook workbook)
        {
            _kernelWorkbookAvailabilityService?.Handle(eventName, workbook, GetKernelHomeForm());
        }

        private bool ShouldAutoShowKernelHomeForEvent(string eventName, Excel.Workbook workbook)
        {
            return _kernelHomeCoordinator.ShouldAutoShowKernelHomeForEvent(eventName, workbook);
        }

        private void HandleExternalWorkbookDetected(Excel.Workbook workbook, string eventName)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.HandleExternalWorkbookDetected(
                _externalWorkbookDetectionService,
                workbook,
                eventName,
                GetKernelHomeForm());
        }

        internal void SuppressUpcomingKernelHomeDisplay(string reason, bool suppressOnOpen, bool suppressOnActivate)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen, suppressOnActivate);
        }

        internal bool ShouldSuppressKernelHomeDisplay(string eventName)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null && _kernelHomeCasePaneSuppressionCoordinator.ShouldSuppressKernelHomeDisplay(eventName);
        }

        internal void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.SuppressUpcomingCasePaneActivationRefresh(workbookFullName, reason);
        }

        internal void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.BeginCaseWorkbookActivateProtection(workbook, window, reason);
        }

        internal bool ShouldIgnoreWorkbookActivateDuringCaseProtection(Excel.Workbook workbook)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldIgnoreWorkbookActivateDuringProtection(workbook);
        }

        internal bool ShouldIgnoreWindowActivateDuringCaseProtection(Excel.Workbook workbook, Excel.Window window)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldIgnoreWindowActivateDuringProtection(workbook, window);
        }

        internal bool ShouldIgnoreTaskPaneRefreshDuringCaseProtection(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldIgnoreTaskPaneRefreshDuringProtection(reason, workbook, window);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneManager != null
                && _taskPaneManager.HasVisibleCasePaneForWorkbookWindow(workbook, window);
        }

        private bool ShouldSuppressCasePaneRefresh(string eventName, Excel.Workbook workbook)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldSuppressCasePaneRefresh(eventName, workbook);
        }

        internal bool IsKernelHomeSuppressionActive(string eventName, bool consume)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.IsKernelHomeSuppressionActive(eventName, consume);
        }

        private Excel.Window ResolveWorkbookPaneWindow(Excel.Workbook workbook, string reason, bool activateWorkbook)
        {
            return _taskPaneRefreshOrchestrationService.ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook);
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        private WindowActivateTaskPaneTriggerFacts CaptureWindowActivateTaskPaneTriggerFacts(
            Excel.Workbook workbook,
            Excel.Window window,
            string captureOwner)
        {
            return new WindowActivateTaskPaneTriggerFacts(
                workbook,
                window,
                FormatWorkbookDescriptor(workbook),
                FormatWindowDescriptor(window),
                FormatActiveExcelState(),
                SafeWorkbookFullName(workbook),
                SafeWindowHwnd(window),
                captureOwner);
        }

        private string FormatActiveExcelState()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            Excel.Window activeWindow = _excelInteropService == null ? null : _excelInteropService.GetActiveWindow();
            return "activeWorkbook=" + FormatWorkbookDescriptor(activeWorkbook) + ",activeWindow=" + FormatWindowDescriptor(activeWindow);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private string SafeWorkbookName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookName(workbook);
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\"";
        }

        private void EnsureKernelFlickerTraceForWorkbookOpen(Excel.Workbook workbook)
        {
            if (!IsKernelWorkbookSafe(workbook))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(KernelFlickerTraceContext.CurrentTraceId))
            {
                return;
            }

            KernelFlickerTraceContext.BeginNewTrace();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=trace-begin trigger=WorkbookOpenKernelDetection workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
        }

        private bool IsKernelWorkbookSafe(Excel.Workbook workbook)
        {
            try
            {
                return workbook != null && IsKernelWorkbook(workbook);
            }
            catch
            {
                return false;
            }
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
                return Convert.ToString(lateBoundWindow.Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
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





