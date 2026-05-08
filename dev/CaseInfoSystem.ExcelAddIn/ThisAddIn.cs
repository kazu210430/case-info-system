using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
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
        private KernelHomeForm _kernelHomeForm;
        private KernelCaseInteractionState _kernelCaseInteractionState;
        private int _taskPaneRefreshSuppressionCount;
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

            // 順序維持: event hook の後に startup 時の HOME 表示判定と初回 pane refresh を行う。
            _logger.Info("ThisAddIn_Startup fired.");
            HookApplicationEvents();
            TryShowKernelHomeFormOnStartup();
            RefreshTaskPane("Startup", null, null);
        }

        private void InitializeStartupDiagnostics()
        {
            _logger = new Logger(ExcelAddInTraceLogWriter.Write);
            ExcelProcessLaunchContextTracer.Trace(_logger);
            AddInDeploymentDiagnosticsTracer.Trace(_logger);
        }

        private AddInCompositionRoot CreateStartupCompositionRoot()
        {
            // Composition Root から VSTO 境界で使う依存と delegate を受け取る。
            return new AddInCompositionRoot(
                this,
                Application,
                _logger,
                // UI / pane
                ResolveWorkbookPaneWindow,
                TryRefreshTaskPane,
                IsTaskPaneRefreshSucceeded,
                () => _kernelHomeForm,
                () => _taskPaneRefreshSuppressionCount,
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
                TaskPaneRefreshOrchestrationService.PendingPaneRefreshMaxAttempts,
                KernelSheetCommandSheetCodeName,
                KernelSheetCommandCellAddress);
        }

        private void ApplyCompositionRoot(AddInCompositionRoot compositionRoot)
        {
            // 基盤
            _excelInteropService = compositionRoot.ExcelInteropService;
            _workbookRoleResolver = compositionRoot.WorkbookRoleResolver;
            _caseWorkbookOpenStrategy = compositionRoot.CaseWorkbookOpenStrategy;

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
            _kernelUserDataReflectionService = compositionRoot.KernelUserDataReflectionService;
            _workbookRibbonCommandService = compositionRoot.WorkbookRibbonCommandService;
            _workbookCaseTaskPaneRefreshCommandService = compositionRoot.WorkbookCaseTaskPaneRefreshCommandService;
            _workbookResetCommandService = compositionRoot.WorkbookResetCommandService;
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

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_logger != null)
            {
                UnhookApplicationEvents();
                _taskPaneRefreshOrchestrationService?.StopPendingPaneRefreshTimer();
                if (_kernelHomeForm != null && !_kernelHomeForm.IsDisposed)
                {
                    _kernelHomeForm.Close();
                    _kernelHomeForm = null;
                }

                if (_taskPaneManager != null)
                {
                    _taskPaneManager.DisposeAll();
                }

                StopWordWarmupTimer();
                _caseWorkbookOpenStrategy?.ShutdownHiddenApplicationCache();

                _logger.Info("ThisAddIn_Shutdown fired.");
                return;
            }

            ExcelAddInTraceLogWriter.Write("ThisAddIn_Shutdown fired.");
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

        // Excel application event handler
        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
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

            if (_kernelHomeForm != null && !_kernelHomeForm.IsDisposed && !_kernelHomeForm.Visible)
            {
                try
                {
                    _kernelHomeForm.PrepareForSilentDispose();
                    _kernelHomeForm.Dispose();
                }
                catch (Exception ex)
                {
                    _logger?.Error("KernelHomeForm dispose before recreation failed.", ex);
                }
                finally
                {
                    _kernelHomeForm = null;
                }
            }

            if (_kernelHomeForm == null || _kernelHomeForm.IsDisposed)
            {
                if (clearBindingOnNewSession)
                {
                    _kernelWorkbookService?.ClearHomeWorkbookBinding("ThisAddIn.ShowKernelHomePlaceholder.NewSession");
                }

                _kernelHomeForm = new KernelHomeForm(_kernelWorkbookService, _kernelCaseCreationCommandService, _logger);
            }

            _taskPaneManager?.HideKernelPanes();
            _kernelHomeForm.ReloadSettings();
            _kernelHomeForm.Invalidate(true);
            _kernelHomeForm.Update();

            TraceRuntimeExecutionObservation("ShowKernelHomePlaceholder");
            _kernelWorkbookService.PrepareForHomeDisplayFromSheet();
            _kernelWorkbookService.EnsureHomeDisplayHidden("ThisAddIn.ShowKernelHomePlaceholder.BeforeShow");

            if (!_kernelHomeForm.Visible)
            {
                _kernelHomeForm.Show();
            }

            _kernelHomeForm.Activate();
            _kernelHomeForm.BringToFront();
        }

        private void TraceRuntimeExecutionObservation(string reason)
        {
            if (_logger == null)
            {
                return;
            }

            try
            {
                Assembly executingAssembly = Assembly.GetExecutingAssembly();
                string assemblyLocation = executingAssembly.Location ?? string.Empty;
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory ?? string.Empty;
                string primaryLogPath = ExcelAddInTraceLogWriter.GetPrimaryTraceLogPath();
                string fallbackLogPath = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.ExcelAddIn", "CaseInfoSystem.ExcelAddIn_trace.log");
                string processId = Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture);
                string assemblyLastWriteTimeUtc = SafeGetAssemblyLastWriteTimeUtc(assemblyLocation);
                string assemblyFileSize = SafeGetAssemblyFileSize(assemblyLocation);
                string assemblySha256 = SafeComputeAssemblySha256(assemblyLocation);
                string assemblyFileVersion = SafeGetAssemblyFileVersion(assemblyLocation);
                string assemblyVersion = SafeGetAssemblyVersion(executingAssembly);
                string assemblyInformationalVersion = SafeGetAssemblyInformationalVersion(executingAssembly);
                string assemblyBuildMarker = SafeGetAssemblyBuildMarker(executingAssembly);

                _logger.Info(
                    "Runtime execution observed. reason=" + (reason ?? string.Empty)
                    + ", assemblyLocation=" + assemblyLocation
                    + ", assemblyLastWriteTimeUtc=" + assemblyLastWriteTimeUtc
                    + ", assemblyFileSize=" + assemblyFileSize
                    + ", assemblySha256=" + assemblySha256
                    + ", assemblyFileVersion=" + assemblyFileVersion
                    + ", assemblyVersion=" + assemblyVersion
                    + ", assemblyInformationalVersion=" + assemblyInformationalVersion
                    + ", assemblyBuildMarker=" + assemblyBuildMarker
                    + ", appDomainBaseDirectory=" + baseDirectory
                    + ", primaryLogPath=" + primaryLogPath
                    + ", fallbackLogPath=" + fallbackLogPath
                    + ", pid=" + processId);
            }
            catch (Exception ex)
            {
                _logger.Error("Runtime execution observation failed. reason=" + (reason ?? string.Empty), ex);
            }
        }

        private static string SafeGetAssemblyLastWriteTimeUtc(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                return new FileInfo(path).LastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyFileSize(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                return new FileInfo(path).Length.ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeComputeAssemblySha256(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                using (FileStream stream = File.OpenRead(path))
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hash = sha256.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", string.Empty);
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyFileVersion(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return string.Empty;
                }

                return FileVersionInfo.GetVersionInfo(path).FileVersion ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyVersion(Assembly assembly)
        {
            try
            {
                return assembly == null
                    ? string.Empty
                    : (assembly.GetName().Version == null ? string.Empty : assembly.GetName().Version.ToString());
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyInformationalVersion(Assembly assembly)
        {
            try
            {
                if (assembly == null)
                {
                    return string.Empty;
                }

                object[] attributes = assembly.GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false);
                if (attributes == null || attributes.Length == 0)
                {
                    return string.Empty;
                }

                AssemblyInformationalVersionAttribute attribute = attributes[0] as AssemblyInformationalVersionAttribute;
                return attribute == null ? string.Empty : attribute.InformationalVersion ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeGetAssemblyBuildMarker(Assembly assembly)
        {
            try
            {
                return assembly == null ? string.Empty : assembly.ManifestModule.ModuleVersionId.ToString("D");
            }
            catch
            {
                return string.Empty;
            }
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
                RunWithScreenUpdatingSuspended(performTransition);
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

        // UI 更新抑止 / 対象解決 / suppression 状態
        internal void RunWithScreenUpdatingSuspended(Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            bool previousScreenUpdating = true;
            try
            {
                previousScreenUpdating = Application.ScreenUpdating;
                Application.ScreenUpdating = false;
                action();
            }
            finally
            {
                try
                {
                    Application.ScreenUpdating = previousScreenUpdating;
                }
                catch
                {
                    // 例外処理: 描画復帰失敗でも業務処理完了を優先する。
                }
            }
        }

        internal IDisposable SuppressTaskPaneRefresh(string reason)
        {
            _taskPaneRefreshSuppressionCount++;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=suppress-enter reason="
                + (reason ?? string.Empty)
                + ", suppressionCount="
                + _taskPaneRefreshSuppressionCount.ToString(CultureInfo.InvariantCulture)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "Task pane refresh suppression entered. reason="
                + (reason ?? string.Empty)
                + ", suppressionCount="
                + _taskPaneRefreshSuppressionCount.ToString());
            return new DelegateDisposable(() =>
            {
                if (_taskPaneRefreshSuppressionCount > 0)
                {
                    _taskPaneRefreshSuppressionCount--;
                }

                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=ThisAddIn action=suppress-exit reason="
                    + (reason ?? string.Empty)
                    + ", suppressionCount="
                    + _taskPaneRefreshSuppressionCount.ToString(CultureInfo.InvariantCulture)
                    + ", activeState="
                    + FormatActiveExcelState());
                _logger?.Info(
                    "Task pane refresh suppression exited. reason="
                    + (reason ?? string.Empty)
                    + ", suppressionCount="
                    + _taskPaneRefreshSuppressionCount.ToString());
            });
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
            if (_kernelHomeForm == null || _kernelHomeForm.IsDisposed || !_kernelHomeForm.Visible)
            {
                return;
            }

            try
            {
                _kernelHomeForm.Hide();
            }
            catch (Exception ex)
            {
                _logger?.Error("HideKernelHomePlaceholder failed.", ex);
            }
        }

        private void TryShowKernelHomeFormOnStartup()
        {
            bool shouldShow = _kernelWorkbookService != null && _kernelWorkbookService.ShouldShowHomeOnStartup();
            _logger.Info("TryShowKernelHomeFormOnStartup shouldShow=" + shouldShow + ", " + (_kernelWorkbookService == null ? string.Empty : _kernelWorkbookService.DescribeStartupState()));
            if (!shouldShow)
            {
                return;
            }

            _kernelWorkbookService?.ClearHomeWorkbookBinding("ThisAddIn.TryShowKernelHomeFormOnStartup");
            ShowKernelHomePlaceholder();
        }

        // Kernel workbook 到達後の UI 反映責務は ThisAddIn に残し、判定自体は coordinator へ委譲する。
        internal void HandleKernelWorkbookBecameAvailable(string eventName, Excel.Workbook workbook)
        {
            _kernelWorkbookAvailabilityService?.Handle(eventName, workbook, _kernelHomeForm);
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
                _kernelHomeForm);
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

        private sealed class DelegateDisposable : IDisposable
        {
            private readonly Action _disposeAction;
            private bool _disposed;

            internal DelegateDisposable(Action disposeAction)
            {
                _disposeAction = disposeAction ?? throw new ArgumentNullException(nameof(disposeAction));
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _disposeAction();
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





