using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
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
        private const string SystemRootFolderName = "案件情報System";
        private const string TraceLogFileName = "CaseInfoSystem.ExcelAddIn_trace.log";
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private static readonly bool DisableSheetActivateForFreezeIsolation = false;
        private static readonly bool DisableSheetSelectionChangeForFreezeIsolation = false;
        private static readonly bool DisableSheetChangeForFreezeIsolation = false;
        private static readonly bool DisableCaseWordWarmupForFreezeIsolation = true;
        private const string KernelSheetCommandSheetCodeName = "shCaseList";
        private const string KernelSheetCommandCellAddress = "AT1";
        private const string ProductTitle = "案件情報System";
        private const int WordWarmupDelayMs = 1500;
        private static readonly Encoding TraceLogEncoding = new UTF8Encoding(false);

        // 基盤
        private Logger _logger;
        private ExcelInteropService _excelInteropService;
        private ExcelWindowRecoveryService _excelWindowRecoveryService;
        private WorkbookRoleResolver _workbookRoleResolver;
        private NavigationService _navigationService;
        private WorkbookSessionService _workbookSessionService;
        private TransientPaneSuppressionService _transientPaneSuppressionService;

        // 文書実行
        private CaseContextFactory _caseContextFactory;
        private DocumentCommandService _documentCommandService;
        private DocumentNamePromptService _documentNamePromptService;
        private DocumentExecutionModeService _documentExecutionModeService;
        private WordInteropService _wordInteropService;

        // workbook ライフサイクル
        private KernelWorkbookService _kernelWorkbookService;
        private KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private CaseWorkbookLifecycleService _caseWorkbookLifecycleService;
        private AccountingWorkbookLifecycleService _accountingWorkbookLifecycleService;
        private AccountingSheetControlService _accountingSheetControlService;
        private WorkbookClipboardPreservationService _workbookClipboardPreservationService;

        // Kernel 操作
        private KernelCaseCreationCommandService _kernelCaseCreationCommandService;
        private KernelUserDataReflectionService _kernelUserDataReflectionService;
        private KernelCommandService _kernelCommandService;
        private WorkbookRibbonCommandService _workbookRibbonCommandService;
        private WorkbookCaseTaskPaneRefreshCommandService _workbookCaseTaskPaneRefreshCommandService;
        private WorkbookResetCommandService _workbookResetCommandService;
        private KernelSheetCommandTriggerService _kernelSheetCommandTriggerService;

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
            _logger = new Logger(LogTrace);
            TraceProcessLaunchContext();

            // Composition Root から VSTO 境界で使う依存と delegate を受け取る。
            var compositionRoot = new AddInCompositionRoot(
                this,
                Application,
                _logger,
                // UI / pane
                ResolveWorkbookPaneWindow,
                IsTaskPaneRefreshSucceeded,
                () => _kernelHomeForm,
                () => _taskPaneRefreshSuppressionCount,
                // Kernel HOME / sheet command
                ShowKernelHomeFromKernelCommand,
                ClearKernelSheetCommandCell,
                ReleaseComObject,
                ShowKernelHomePlaceholderWithExternalWorkbookSuppression,
                HandleExternalWorkbookDetected,
                ShouldSuppressCasePaneRefresh,
                // 非同期 UI 更新
                RefreshTaskPane,
                ScheduleWordWarmup,
                TaskPaneRefreshOrchestrationService.PendingPaneRefreshMaxAttempts,
                KernelSheetCommandSheetCodeName,
                KernelSheetCommandCellAddress);
            compositionRoot.Compose();
            ApplyCompositionRoot(compositionRoot);

            _logger.Info("ThisAddIn_Startup fired.");
            HookApplicationEvents();
            TryShowKernelHomeFormOnStartup();
            RefreshTaskPane("Startup", null, null);
        }

        private void ApplyCompositionRoot(AddInCompositionRoot compositionRoot)
        {
            // 基盤
            _excelInteropService = compositionRoot.ExcelInteropService;
            _excelWindowRecoveryService = compositionRoot.ExcelWindowRecoveryService;
            _workbookRoleResolver = compositionRoot.WorkbookRoleResolver;
            _navigationService = compositionRoot.NavigationService;
            _workbookSessionService = compositionRoot.WorkbookSessionService;
            _transientPaneSuppressionService = compositionRoot.TransientPaneSuppressionService;

            // 文書実行
            _caseContextFactory = compositionRoot.CaseContextFactory;
            _documentCommandService = compositionRoot.DocumentCommandService;
            _documentNamePromptService = compositionRoot.DocumentNamePromptService;
            _documentExecutionModeService = compositionRoot.DocumentExecutionModeService;
            _wordInteropService = compositionRoot.WordInteropService;

            // workbook ライフサイクル
            _kernelWorkbookService = compositionRoot.KernelWorkbookService;
            _kernelWorkbookLifecycleService = compositionRoot.KernelWorkbookLifecycleService;
            _caseWorkbookLifecycleService = compositionRoot.CaseWorkbookLifecycleService;
            _accountingWorkbookLifecycleService = compositionRoot.AccountingWorkbookLifecycleService;
            _accountingSheetControlService = compositionRoot.AccountingSheetControlService;
            _workbookClipboardPreservationService = compositionRoot.WorkbookClipboardPreservationService;

            // Kernel 操作
            _kernelCaseCreationCommandService = compositionRoot.KernelCaseCreationCommandService;
            _kernelUserDataReflectionService = compositionRoot.KernelUserDataReflectionService;
            _kernelCommandService = compositionRoot.KernelCommandService;
            _workbookRibbonCommandService = compositionRoot.WorkbookRibbonCommandService;
            _workbookCaseTaskPaneRefreshCommandService = compositionRoot.WorkbookCaseTaskPaneRefreshCommandService;
            _workbookResetCommandService = compositionRoot.WorkbookResetCommandService;
            _kernelSheetCommandTriggerService = compositionRoot.KernelSheetCommandTriggerService;

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

                _logger.Info("ThisAddIn_Shutdown fired.");
                return;
            }

            LogTrace("ThisAddIn_Shutdown fired.");
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
            ((Excel.AppEvents_Event)Application).WorkbookOpen += Application_WorkbookOpen;
            ((Excel.AppEvents_Event)Application).WorkbookActivate += Application_WorkbookActivate;
            ((Excel.AppEvents_Event)Application).WorkbookBeforeSave += Application_WorkbookBeforeSave;
            ((Excel.AppEvents_Event)Application).WorkbookBeforeClose += Application_WorkbookBeforeClose;
            ((Excel.AppEvents_Event)Application).WindowActivate += Application_WindowActivate;
            if (!DisableSheetActivateForFreezeIsolation)
            {
                ((Excel.AppEvents_Event)Application).SheetActivate += Application_SheetActivate;
            }
            if (!DisableSheetSelectionChangeForFreezeIsolation)
            {
                ((Excel.AppEvents_Event)Application).SheetSelectionChange += Application_SheetSelectionChange;
            }
            if (!DisableSheetChangeForFreezeIsolation)
            {
                ((Excel.AppEvents_Event)Application).SheetChange += Application_SheetChange;
            }
            ((Excel.AppEvents_Event)Application).AfterCalculate += Application_AfterCalculate;
        }

        private void UnhookApplicationEvents()
        {
            ((Excel.AppEvents_Event)Application).WorkbookOpen -= Application_WorkbookOpen;
            ((Excel.AppEvents_Event)Application).WorkbookActivate -= Application_WorkbookActivate;
            ((Excel.AppEvents_Event)Application).WorkbookBeforeSave -= Application_WorkbookBeforeSave;
            ((Excel.AppEvents_Event)Application).WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            ((Excel.AppEvents_Event)Application).WindowActivate -= Application_WindowActivate;
            if (!DisableSheetActivateForFreezeIsolation)
            {
                ((Excel.AppEvents_Event)Application).SheetActivate -= Application_SheetActivate;
            }
            if (!DisableSheetSelectionChangeForFreezeIsolation)
            {
                ((Excel.AppEvents_Event)Application).SheetSelectionChange -= Application_SheetSelectionChange;
            }
            if (!DisableSheetChangeForFreezeIsolation)
            {
                ((Excel.AppEvents_Event)Application).SheetChange -= Application_SheetChange;
            }
            ((Excel.AppEvents_Event)Application).AfterCalculate -= Application_AfterCalculate;
        }

        // Excel application event handler
        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            EnsureKernelFlickerTraceForWorkbookOpen(workbook);
            EventBoundaryGuard.Execute(_logger, nameof(Application_WorkbookOpen), () =>
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=ExcelEventBoundary action=fire event=WorkbookOpen workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveExcelState());
                _logger?.Info("Excel WorkbookOpen fired. workbook=" + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));
                _workbookEventCoordinator.OnWorkbookOpen(workbook);
            });
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_WorkbookActivate), () =>
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=ExcelEventBoundary action=fire event=WorkbookActivate workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveExcelState());
                _logger?.Info("Excel WorkbookActivate fired. workbook=" + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));
                _workbookEventCoordinator.OnWorkbookActivate(workbook);
            });
        }

        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_WindowActivate), () =>
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=ExcelEventBoundary action=fire event=WindowActivate workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", eventWindow="
                    + FormatWindowDescriptor(window)
                    + ", activeState="
                    + FormatActiveExcelState());
                _logger?.Info(
                    "Excel WindowActivate fired. workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook))
                    + ", windowHwnd="
                    + SafeWindowHwnd(window));
                _workbookEventCoordinator.OnWindowActivate(workbook, window);
            });
        }

        // event handler から調停 service へ渡す delegate 入口
        internal void HandleWorkbookOpenEvent(Excel.Workbook workbook)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WorkbookOpen workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info("TaskPane event entry. event=WorkbookOpen, workbook=" + SafeWorkbookFullName(workbook) + ", activeWorkbook=" + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook()) + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));

            // 外部 workbook 検知 -> lifecycle 同期 -> Kernel HOME 判定 -> pane 更新
            HandleExternalWorkbookDetected(workbook, "WorkbookOpen");
            _kernelWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _accountingWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _accountingSheetControlService?.EnsureVstoManagedControls(workbook);
            _caseWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _kernelHomeCoordinator.HandleKernelWorkbookBecameAvailable("WorkbookOpen", workbook);
            RefreshTaskPane("WorkbookOpen", workbook, null);
        }

        internal void HandleWorkbookActivateEvent(Excel.Workbook workbook)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WorkbookActivate workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info("TaskPane event entry. event=WorkbookActivate, workbook=" + SafeWorkbookFullName(workbook) + ", activeWorkbook=" + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook()) + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));

            // 外部 workbook 検知 -> lifecycle 同期 -> Kernel HOME 判定 -> pane 更新
            HandleExternalWorkbookDetected(workbook, "WorkbookActivate");
            _kernelWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _accountingWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _accountingSheetControlService?.EnsureVstoManagedControls(workbook);
            _caseWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _kernelHomeCoordinator.HandleKernelWorkbookBecameAvailable("WorkbookActivate", workbook);
            if (ShouldSuppressCasePaneRefresh("WorkbookActivate", workbook))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=WorkbookEventCoordinator action=suppress-refresh event=WorkbookActivate workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveExcelState());
                return;
            }

            RefreshTaskPane("WorkbookActivate", workbook, null);
        }

        internal void HandleWindowActivateEvent(Excel.Workbook workbook, Excel.Window window)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WindowActivate workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", eventWindow="
                + FormatWindowDescriptor(window)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info("TaskPane event entry. event=WindowActivate, workbook=" + SafeWorkbookFullName(workbook) + ", windowHwnd=" + SafeWindowHwnd(window) + ", activeWorkbook=" + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook()) + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
            _windowActivatePaneHandlingService?.Handle(workbook, window);
        }

        private void Application_SheetActivate(object sh)
        {
            _logger?.Debug("Application_SheetActivate", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetActivate), () =>
            {
                _accountingWorkbookLifecycleService?.HandleSheetActivated(sh);
                _accountingSheetControlService?.HandleSheetActivated(sh);
                _caseWorkbookLifecycleService?.HandleSheetActivated(sh);
                RefreshTaskPane("SheetActivate", null, null);
            });
            _logger?.Debug("Application_SheetActivate", "returned.");
        }

        private void Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            _logger?.Debug("Application_SheetSelectionChange", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetSelectionChange), () =>
            {
                string sheetName = SafeSheetName(sh);
                string targetAddress = SafeRangeAddress(target);
                _logger?.Debug("Application_SheetSelectionChange", "fired. sheet=" + sheetName + ", target=" + targetAddress);
                _accountingSheetControlService?.HandleSheetSelectionChange(sh, target);
            });
            _logger?.Debug("Application_SheetSelectionChange", "returned.");
        }

        private void Application_SheetChange(object sh, Excel.Range target)
        {
            _logger?.Debug("Application_SheetChange", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetChange), () =>
            {
                string sheetName = SafeSheetName(sh);
                string targetAddress = SafeRangeAddress(target);
                _logger?.Debug("Application_SheetChange", "fired. sheet=" + sheetName + ", target=" + targetAddress);
                HandleKernelSheetCommand(sh as Excel.Worksheet, target);
                _caseWorkbookLifecycleService?.HandleSheetChanged((sh as Excel.Worksheet)?.Parent as Excel.Workbook);
                _accountingSheetControlService?.HandleSheetChange(sh, target);
            });
            _logger?.Debug("Application_SheetChange", "returned.");
        }

        private void Application_AfterCalculate()
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_AfterCalculate), () =>
            {
                _logger?.Debug("Application_AfterCalculate", "fired.");
                _accountingSheetControlService?.HandleAfterCalculate(Application);
                _logger?.Debug("Application_AfterCalculate", "after AccountingSheetControlService.HandleAfterCalculate returned.");
            });
            _logger?.Debug("Application_AfterCalculate", "EventBoundaryGuard.Execute returned.");
        }

        private void HandleKernelSheetCommand(Excel.Worksheet worksheet, Excel.Range target)
        {
            _kernelSheetCommandTriggerService?.Handle(worksheet, target);
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

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
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

        // 起動診断ログ
        private void TraceProcessLaunchContext()
        {
            if (_logger == null)
            {
                return;
            }

            try
            {
                Process current = Process.GetCurrentProcess();
                int currentProcessId = current.Id;
                int parentProcessId = TryGetParentProcessId(currentProcessId);
                string parentSummary = GetProcessSummary(parentProcessId);
                string currentCommandLine = TryGetProcessCommandLine(currentProcessId);
                string excelProcesses = BuildExcelProcessSnapshot(currentProcessId);

                _logger.Info(
                    "Process launch context. currentPid="
                    + currentProcessId.ToString(CultureInfo.InvariantCulture)
                    + ", currentName="
                    + SafeProcessName(current)
                    + ", sessionId="
                    + current.SessionId.ToString(CultureInfo.InvariantCulture)
                    + ", startTime="
                    + SafeProcessStartTime(current)
                    + ", parentPid="
                    + parentProcessId.ToString(CultureInfo.InvariantCulture)
                    + ", parent="
                    + parentSummary
                    + ", commandLine="
                    + currentCommandLine
                    + ", excelProcesses=["
                    + excelProcesses
                    + "]");
            }
            catch (Exception ex)
            {
                _logger.Error("TraceProcessLaunchContext failed.", ex);
            }
        }

        private static int TryGetParentProcessId(int processId)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = " + processId.ToString(CultureInfo.InvariantCulture)))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {
                        object parentProcessId = process["ParentProcessId"];
                        if (parentProcessId == null)
                        {
                            return 0;
                        }

                        return Convert.ToInt32(parentProcessId, CultureInfo.InvariantCulture);
                    }
                }
            }
            catch
            {
                return 0;
            }

            return 0;
        }

        private static string TryGetProcessCommandLine(int processId)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + processId.ToString(CultureInfo.InvariantCulture)))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {
                        object commandLine = process["CommandLine"];
                        return commandLine == null ? string.Empty : Convert.ToString(commandLine, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        private static string GetProcessSummary(int processId)
        {
            if (processId <= 0)
            {
                return "(unknown)";
            }

            try
            {
                using (Process process = Process.GetProcessById(processId))
                {
                    return "pid="
                        + process.Id.ToString(CultureInfo.InvariantCulture)
                        + ",name="
                        + SafeProcessName(process)
                        + ",startTime="
                        + SafeProcessStartTime(process);
                }
            }
            catch
            {
                return "pid=" + processId.ToString(CultureInfo.InvariantCulture) + ",name=(unavailable)";
            }
        }

        private static string BuildExcelProcessSnapshot(int currentProcessId)
        {
            var builder = new StringBuilder();

            try
            {
                Process[] processes = Process.GetProcessesByName("EXCEL");
                foreach (Process process in processes)
                {
                    using (process)
                    {
                        if (builder.Length > 0)
                        {
                            builder.Append(" | ");
                        }

                        builder.Append("pid=");
                        builder.Append(process.Id.ToString(CultureInfo.InvariantCulture));
                        builder.Append(",name=");
                        builder.Append(SafeProcessName(process));
                        builder.Append(",startTime=");
                        builder.Append(SafeProcessStartTime(process));
                        builder.Append(",isCurrent=");
                        builder.Append((process.Id == currentProcessId).ToString());
                        builder.Append(",commandLine=");
                        builder.Append(TryGetProcessCommandLine(process.Id));
                    }
                }
            }
            catch
            {
                return builder.ToString();
            }

            return builder.ToString();
        }

        private static string SafeProcessName(Process process)
        {
            try
            {
                return process == null ? string.Empty : process.ProcessName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeProcessStartTime(Process process)
        {
            try
            {
                if (process == null)
                {
                    return string.Empty;
                }

                return process.StartTime.ToString("O", CultureInfo.InvariantCulture);
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
                _logger?.Info(
                    "Excel WorkbookBeforeClose fired. workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook))
                    + ", cancel="
                    + innerCancel.ToString());

                if (innerCancel)
                {
                    return;
                }

                _caseWorkbookLifecycleService?.HandleWorkbookBeforeClose(workbook, ref innerCancel);
                if (innerCancel)
                {
                    return;
                }

                _kernelWorkbookLifecycleService?.HandleWorkbookBeforeClose(workbook, ref innerCancel);
                if (innerCancel)
                {
                    return;
                }

                _workbookClipboardPreservationService?.PreserveCopiedValuesForClosingWorkbook(workbook);
                _accountingWorkbookLifecycleService?.HandleWorkbookBeforeClose(workbook);

                if (_accountingSheetControlService != null)
                {
                    _accountingSheetControlService.RemoveWorkbookState(workbook);
                }

                if (_accountingWorkbookLifecycleService != null)
                {
                    _accountingWorkbookLifecycleService.RemoveWorkbookState(workbook);
                }

                if (_caseWorkbookLifecycleService != null)
                {
                    _caseWorkbookLifecycleService.RemoveWorkbookState(workbook);
                }

                if (_taskPaneManager != null)
                {
                    _taskPaneManager.RemoveWorkbookPanes(workbook);
                }
            }
        }

        // Task pane / HOME 表示の VSTO 境界
        private void RefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
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
                + FormatActiveExcelState());
            TaskPaneRefreshAttemptResult result = TryRefreshTaskPane(reason, workbook, window);
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
                + (result == null ? "null" : result.IsRefreshSucceeded.ToString()));
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

        internal CustomTaskPane CreateTaskPane(Excel.Window window, System.Windows.Forms.UserControl control)
        {
            CustomTaskPane pane = CustomTaskPanes.Add(control, TaskPaneTitle, window);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            return pane;
        }

        internal void RemoveTaskPane(CustomTaskPane pane)
        {
            if (pane == null)
            {
                return;
            }

            CustomTaskPanes.Remove(pane);
        }

        private void LogTrace(string message)
        {
            string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
                + " [PID=" + Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture) + "] CaseInfoSystem: "
                + (message ?? string.Empty);

            try
            {
                string logDirectory = ResolveLogDirectory();
                Directory.CreateDirectory(logDirectory);
                File.AppendAllText(Path.Combine(logDirectory, TraceLogFileName), line + Environment.NewLine, TraceLogEncoding);
            }
            catch
            {
                // ログ書き込み失敗時でも Add-in 起動は継続する。
            }

            try
            {
                string fallbackDirectory = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.ExcelAddIn");
                Directory.CreateDirectory(fallbackDirectory);
                File.AppendAllText(Path.Combine(fallbackDirectory, TraceLogFileName), line + Environment.NewLine, TraceLogEncoding);
            }
            catch
            {
                // フォールバックログ失敗時でも Add-in は停止させない。
            }
        }
        private static string ResolveLogDirectory()
        {
            string userDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(userDocuments, SystemRootFolderName, "logs");
        }

        // Kernel HOME / sheet 遷移
        private void ShowKernelHomePlaceholder()
        {
            _kernelHomeCasePaneSuppressionCoordinator?.ResetKernelHomeExternalCloseRequested();

            if (_kernelHomeForm != null && !_kernelHomeForm.IsDisposed && !_kernelHomeForm.Visible)
            {
                try
                {
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
                _kernelHomeForm = new KernelHomeForm(_kernelWorkbookService, _kernelCaseCreationCommandService, _logger);
            }

            _taskPaneManager?.HideKernelPanes();
            _kernelHomeForm.ReloadSettings();
            _kernelHomeForm.Invalidate(true);
            _kernelHomeForm.Update();

            if (!_kernelHomeForm.Visible)
            {
                _kernelHomeForm.Show();
            }

            _kernelWorkbookService.PrepareForHomeDisplayFromSheet();
            _kernelHomeForm.Activate();
            _kernelHomeForm.BringToFront();
        }

        internal bool ShowKernelSheetAndRefreshPaneFromHome(string sheetCodeName, string reason)
        {
            KernelFlickerTraceContext.BeginNewTrace();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=trace-begin trigger=ShowKernelSheetAndRefreshPaneFromHome traceOriginReason="
                + (reason ?? string.Empty)
                + ", sheetCodeName="
                + (sheetCodeName ?? string.Empty)
                + ", activeState="
                + FormatActiveExcelState());
            // 処理ブロック: 次に来る activate 系イベントに備えて、Kernel HOME 抑止要求を発行する。
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            bool shown = _kernelWorkbookService.ShowSheetByCodeName(sheetCodeName);
            if (!shown)
            {
                return false;
            }

            RefreshTaskPane(reason, _kernelWorkbookService.GetOpenKernelWorkbook(), null);
            return true;
        }

        internal bool ShowKernelSheetAndRefreshPane(string sheetCodeName, string reason)
        {
            KernelFlickerTraceContext.BeginNewTrace();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=trace-begin trigger=ShowKernelSheetAndRefreshPane traceOriginReason="
                + (reason ?? string.Empty)
                + ", sheetCodeName="
                + (sheetCodeName ?? string.Empty)
                + ", activeState="
                + FormatActiveExcelState());
            // 処理ブロック: 遷移開始前の時点観測として、開始ログへ出す workbook 状態を記録する。
            Excel.Workbook activeWorkbookBefore = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            _logger?.Info(
                "ShowKernelSheetAndRefreshPane started. reason="
                + (reason ?? string.Empty)
                + ", sheetCodeName="
                + (sheetCodeName ?? string.Empty)
                + ", activeWorkbookBefore="
                + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(activeWorkbookBefore)));
            // 処理ブロック: 次に来る activate 系イベントに備えて、Kernel HOME 抑止要求を発行する。
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            _logger?.Info("[Transition] suppression requested. reason=" + reason);
            // 処理ブロック: sheet 表示前の内部 cleanup として、表示中の HOME UI を退避する。
            HideKernelHomePlaceholder();
            // 処理ブロック: 表示実行そのものではなく、対象 sheet の表示要求を発行し、その結果を受けて続行可否を判定する。
            bool shown = _kernelWorkbookService.ShowSheetByCodeName(sheetCodeName);
            _logger?.Info("[Transition] sheet shown=" + shown + ", sheet=" + sheetCodeName);
            if (!shown)
            {
                // 処理ブロック: sheet 表示失敗時の中断観測として、失敗理由をログへ記録して終了する。
                _logger?.Info(
                    "ShowKernelSheetAndRefreshPane aborted because target sheet could not be shown. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            // 処理ブロック: pane 同期の実行ではなく、表示後の pane 同期要求を発行する。
            Excel.Workbook kernelWorkbook = _kernelWorkbookService.GetOpenKernelWorkbook();
            _logger?.Info("[Transition] pane refresh requested.");
            RefreshTaskPane(reason, kernelWorkbook, null);
            // 処理ブロック: 遷移完了後の時点観測として、完了ログへ出す workbook 状態を記録する。
            Excel.Workbook activeWorkbookAfter = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            _logger?.Info(
                "ShowKernelSheetAndRefreshPane completed. reason="
                + (reason ?? string.Empty)
                + ", kernelWorkbook="
                + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(kernelWorkbook))
                + ", activeWorkbookAfter="
                + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(activeWorkbookAfter)));
            return true;
        }

        internal void ShowKernelHomeFromAutomation()
        {
            _logger?.Info("Kernel home requested from COM automation.");
            if (_logger == null)
            {
                LogTrace("Kernel home requested from COM automation.");
            }
            ShowKernelHomePlaceholderWithExternalWorkbookSuppression("KernelAutomationService.ShowHome");
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

            LogTrace((message ?? string.Empty) + " exception=" + (ex == null ? string.Empty : ex.ToString()));
        }

        public void ShowKernelHomeFromSheet()
        {
            ShowKernelHomeFromAutomation();
        }

        // Kernel HOME の明示表示直後だけ activate 系イベントの自動クローズを抑止する。
        private void ShowKernelHomePlaceholderWithExternalWorkbookSuppression(string reason)
        {
            // 処理ブロック: HOME を明示表示した直後だけ WorkbookActivate/WindowActivate を抑止する。
            // ここでは将来の activate 系イベントに効く Kernel HOME 抑止要求を発行する。
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            ShowKernelHomePlaceholder();
        }

        public void ReflectKernelUserDataToAccountingSet()
        {
            _logger?.Info("Kernel user data reflection requested from COM automation. target=AccountingSet");
            _kernelUserDataReflectionService?.ReflectToAccountingSetOnly();
        }

        public void ReflectKernelUserDataToBaseHome()
        {
            _logger?.Info("Kernel user data reflection requested from COM automation. target=BaseHome");
            _kernelUserDataReflectionService?.ReflectToBaseHomeOnly();
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
                LogTrace("COM automation service requested before startup.");
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

            if (_documentExecutionModeService == null || !_documentExecutionModeService.CanAttemptVstoExecution())
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





