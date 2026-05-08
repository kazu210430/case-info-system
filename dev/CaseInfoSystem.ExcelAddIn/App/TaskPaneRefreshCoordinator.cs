using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshCoordinator
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly TaskPaneManager _taskPaneManager;
        private readonly ExcelWindowRecoveryService _excelWindowRecoveryService;
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Action _scheduleWordWarmup;
        private readonly ICasePaneHostBridge _casePaneHostBridge;
        private int _kernelFlickerTraceCoordinatorAttemptSequence;

        internal TaskPaneRefreshCoordinator(
            WorkbookSessionService workbookSessionService,
            TaskPaneManager taskPaneManager,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            Logger logger,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Action scheduleWordWarmup,
            ICasePaneHostBridge casePaneHostBridge)
        {
            _workbookSessionService = workbookSessionService;
            _taskPaneManager = taskPaneManager;
            _excelWindowRecoveryService = excelWindowRecoveryService;
            _logger = logger;
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow;
            _scheduleWordWarmup = scheduleWordWarmup;
            _casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException(nameof(casePaneHostBridge));
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window, KernelHomeForm kernelHomeForm, int taskPaneRefreshSuppressionCount)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            int coordinatorAttemptId = ++_kernelFlickerTraceCoordinatorAttemptSequence;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshCoordinator action=start coordinatorAttemptId="
                + coordinatorAttemptId.ToString()
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", kernelHomeVisible="
                + (kernelHomeForm != null && !kernelHomeForm.IsDisposed && kernelHomeForm.Visible).ToString()
                + ", suppressionCount="
                + taskPaneRefreshSuppressionCount.ToString());
            if (!CanExecuteTaskPaneRefresh(reason, stopwatch, taskPaneRefreshSuppressionCount))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=end coordinatorAttemptId="
                    + coordinatorAttemptId.ToString()
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", result=Skipped");
                return TaskPaneRefreshAttemptResult.Skipped();
            }

            if (TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=end coordinatorAttemptId="
                    + coordinatorAttemptId.ToString()
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", result=SkippedWorkbookOpenWindowDependentRefresh"
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", inputWindow="
                    + FormatWindowDescriptor(window));
                return TaskPaneRefreshAttemptResult.Skipped();
            }

            // recovery は単なる UI 修復ではなく、後続の context 解決の前提調整。
            // ActiveWindow / 可視 window / UI 状態を「解決可能な状態」に整える段階であり、
            // ここでは対象 window も context もまだ決定しない。
            bool preContextRecoveryAttempted = false;
            bool? preContextRecoverySucceeded = null;
            if ((kernelHomeForm == null || kernelHomeForm.IsDisposed || !kernelHomeForm.Visible)
                && _excelWindowRecoveryService != null)
            {
                preContextRecoveryAttempted = true;
                if (workbook != null)
                {
                    preContextRecoverySucceeded = _excelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(workbook, "TryRefreshTaskPane." + (reason ?? string.Empty), bringToFront: false);
                }
                else
                {
                    preContextRecoverySucceeded = _excelWindowRecoveryService.TryRecoverActiveWorkbookWindowWithoutShowing("TryRefreshTaskPane." + (reason ?? string.Empty), bringToFront: false);
                }
            }

            // recovery の次に、workbook 指定時の pane 対象 window をここで確定させる。
            // この段階で「どの window を対象にするか」が決まり、
            // 後続の context 生成はこの確定済み window に依存する。
            window = EnsurePaneWindowForWorkbook(workbook, window, reason, stopwatch);

            // ここでは pane 更新に使う実行 context を生成する。
            // context.Window は pane の表示先 / hide 先の正本だが、
            // この段階ではまだ表示採否や hide 判断は行わない。
            WorkbookContext context = workbook == null
                ? _workbookSessionService.ResolveActiveContext(reason)
                : _workbookSessionService.ResolveContext(workbook, window, reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshCoordinator action=context-resolved coordinatorAttemptId="
                + coordinatorAttemptId.ToString()
                + ", reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", resolvedWindow="
                + FormatWindowDescriptor(window)
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString());

            // ここでは生成済み context を pane 対象として採用するかを調停する。
            // 対象外なら context を使わず、必要に応じて hide もここで判断する。
            // 生成と受理は分離されており、context 生成 → 受理判定の順で直列に進む。
            if (!TryAcceptTaskPaneContext(context, window, kernelHomeForm))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=end coordinatorAttemptId="
                    + coordinatorAttemptId.ToString()
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", result=ContextRejected"
                    + ", context="
                    + FormatContextDescriptor(context)
                    + ", resolvedWindow="
                    + FormatWindowDescriptor(window));
                return TaskPaneRefreshAttemptResult.ContextRejected(
                    preContextRecoveryAttempted,
                    preContextRecoverySucceeded);
            }

            // 受理された context を使って pane UI へ反映し、
            // CASE 成功時の warmup もこの最終段でまとめて扱う。
            TaskPaneHostFlowResult hostFlowResult = TryRefreshPaneAndScheduleWarmup(context, reason, stopwatch);
            bool refreshed = hostFlowResult.IsShown;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshCoordinator action=end coordinatorAttemptId="
                + coordinatorAttemptId.ToString()
                + ", reason="
                + (reason ?? string.Empty)
                + ", result="
                + (refreshed ? "Succeeded" : "Failed")
                + ", context="
                + FormatContextDescriptor(context)
                + ", resolvedWindow="
                + FormatWindowDescriptor(window)
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString());
            return refreshed
                ? TaskPaneRefreshAttemptResult.RefreshCompletedPendingForeground(
                    context,
                    workbook,
                    window,
                    _excelWindowRecoveryService != null,
                    "refreshCompleted",
                    hostFlowResult.PaneVisibleSource,
                    hostFlowResult.SnapshotBuildResult,
                    preContextRecoveryAttempted,
                    preContextRecoverySucceeded)
                : TaskPaneRefreshAttemptResult.Failed(
                    hostFlowResult.SnapshotBuildResult,
                    preContextRecoveryAttempted,
                    preContextRecoverySucceeded);
        }

        private bool CanExecuteTaskPaneRefresh(string reason, Stopwatch stopwatch, int taskPaneRefreshSuppressionCount)
        {
            if (_workbookSessionService == null || _taskPaneManager == null)
            {
                _logger?.Warn(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=skip reason="
                    + (reason ?? string.Empty)
                    + ", result=MissingDependency");
                return false;
            }

            if (taskPaneRefreshSuppressionCount > 0)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=skip reason="
                    + (reason ?? string.Empty)
                    + ", result=Suppressed"
                    + ", suppressionCount="
                    + taskPaneRefreshSuppressionCount.ToString());
                _logger?.Info(
                    "TryRefreshTaskPane suppressed. reason="
                    + (reason ?? string.Empty)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString()
                    + ", suppressionCount="
                    + taskPaneRefreshSuppressionCount.ToString());
                return false;
            }

            return true;
        }

        private Excel.Window EnsurePaneWindowForWorkbook(Excel.Workbook workbook, Excel.Window window, string reason, Stopwatch stopwatch)
        {
            // workbook は既に特定できているが、pane 操作に使う window が未確定な場合は、
            // その workbook 専用に window を解決する。
            // ここで補完した window は、後続の context 解決と hide/show 対象 window の決定に使われる。
            if (workbook != null && window == null)
            {
                window = _resolveWorkbookPaneWindow(workbook, reason, false);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=ensure-window reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", resolvedWindow="
                    + FormatWindowDescriptor(window)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString());
            }

            return window;
        }

        private bool TryAcceptTaskPaneContext(WorkbookContext context, Excel.Window window, KernelHomeForm kernelHomeForm)
        {
            if (kernelHomeForm != null && !kernelHomeForm.IsDisposed && kernelHomeForm.Visible)
            {
                if (context != null && context.Role == WorkbookRole.Kernel)
                {
                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneRefreshCoordinator action=context-rejected reason=KernelHomeVisibleWithKernelContext"
                        + ", context="
                        + FormatContextDescriptor(context));
                    _taskPaneManager.HideKernelPanes();
                    return false;
                }
            }

            if (!_workbookSessionService.ShouldHandleContext(context))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshCoordinator action=context-rejected reason=ShouldHandleContextFalse"
                    + ", context="
                    + FormatContextDescriptor(context)
                    + ", fallbackWindow="
                    + FormatWindowDescriptor(window));
                if (context != null)
                {
                    _taskPaneManager.HidePaneForWindow(context.Window);
                }
                else if (window != null)
                {
                    _taskPaneManager.HidePaneForWindow(window);
                }

                return false;
            }

            return true;
        }

        private TaskPaneHostFlowResult TryRefreshPaneAndScheduleWarmup(WorkbookContext context, string reason, Stopwatch stopwatch)
        {
            string observedWorkbookPath = ResolveObservedWorkbookPath(context, null);
            NewCaseVisibilityObservation.Log(
                _logger,
                null,
                null,
                context == null ? null : context.Workbook,
                context == null ? null : context.Window,
                "taskpane-refresh-started",
                "TaskPaneRefreshCoordinator.TryRefreshPaneAndScheduleWarmup",
                observedWorkbookPath,
                "reason=" + (reason ?? string.Empty) + ",refreshSource=" + (reason ?? string.Empty));
            TaskPaneHostFlowResult hostFlowResult = _taskPaneManager.RefreshPaneWithOutcome(context, reason);
            bool refreshed = hostFlowResult.IsShown;
            NewCaseDefaultTimingLogHelper.LogTaskPaneReadyWaitToRefreshCompleted(
                _logger,
                context == null ? string.Empty : SafeWorkbookFullName(context.Workbook, context.WorkbookFullName),
                reason,
                refreshed,
                hostFlowResult.PaneVisibleBasis);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshCoordinator action=refresh-pane-complete reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", refreshed="
                + refreshed.ToString()
                + ", refreshSource="
                + (reason ?? string.Empty)
                + ", paneVisibleBasis="
                + hostFlowResult.PaneVisibleBasis
                + ", scheduleWarmup="
                + (refreshed && context != null && context.Role == WorkbookRole.Case).ToString()
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString()
                + FormatObservationCorrelationFields(context, null));
            NewCaseVisibilityObservation.Log(
                _logger,
                null,
                null,
                context == null ? null : context.Workbook,
                context == null ? null : context.Window,
                "taskpane-refresh-completed",
                "TaskPaneRefreshCoordinator.TryRefreshPaneAndScheduleWarmup",
                observedWorkbookPath,
                "reason=" + (reason ?? string.Empty) + ",refreshSource=" + (reason ?? string.Empty) + ",refreshed=" + refreshed.ToString() + ",paneVisibleBasis=" + hostFlowResult.PaneVisibleBasis);
            if (refreshed && context != null && context.Role == WorkbookRole.Case)
            {
                _scheduleWordWarmup();
            }

            return hostFlowResult;
        }

        internal ForegroundGuaranteeExecutionResult ExecuteFinalForegroundGuaranteeRecovery(WorkbookContext context, Excel.Workbook workbook, string reason)
        {
            if (_excelWindowRecoveryService == null)
            {
                return ForegroundGuaranteeExecutionResult.NotAttempted();
            }

            Stopwatch stopwatch2 = Stopwatch.StartNew();
            bool recovered = workbook != null
                ? _excelWindowRecoveryService.TryRecoverWorkbookWindow(workbook, "TryRefreshTaskPane.PostRefresh." + (reason ?? string.Empty), bringToFront: true)
                : _excelWindowRecoveryService.TryRecoverActiveWorkbookWindow("TryRefreshTaskPane.PostRefresh." + (reason ?? string.Empty), bringToFront: true);
            return new ForegroundGuaranteeExecutionResult(
                executionAttempted: true,
                recovered: recovered,
                elapsedMilliseconds: stopwatch2.ElapsedMilliseconds);
        }

        internal void BeginPostForegroundProtection(WorkbookContext context, Excel.Workbook workbook, string reason, long elapsedMilliseconds)
        {
            Excel.Workbook protectedWorkbook = context == null ? workbook : context.Workbook;
            Excel.Window protectedWindow = context == null ? null : context.Window;
            bool protectionStartRequested = context != null && context.Role == WorkbookRole.Case && protectedWorkbook != null && protectedWindow != null;
            string protectionSkipReason = string.Empty;
            if (context == null || context.Role != WorkbookRole.Case)
            {
                protectionSkipReason = "context.Role!=Case";
            }
            else if (protectedWindow == null)
            {
                protectionSkipReason = "protectedWindow=null";
            }
            else if (protectedWorkbook == null)
            {
                protectionSkipReason = "protectedWorkbook=null";
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshCoordinator action=protection-decision reason="
                + (reason ?? string.Empty)
                + ", contextRole="
                + (context == null ? "null" : context.Role.ToString())
                + ", protectedWorkbookPresent="
                + (protectedWorkbook != null).ToString()
                + ", protectedWindowPresent="
                + (protectedWindow != null).ToString()
                + ", protectionStartRequested="
                + protectionStartRequested.ToString()
                + ", protectionSkipped="
                + (!protectionStartRequested).ToString()
                + ", protectionSkipReason="
                + protectionSkipReason
                + ", elapsedMs="
                + elapsedMilliseconds.ToString()
                + FormatObservationCorrelationFields(context, protectedWorkbook));
            if (protectionStartRequested)
            {
                _casePaneHostBridge.BeginCaseWorkbookActivateProtection(
                    protectedWorkbook,
                    protectedWindow,
                    "TryRefreshTaskPane.PostRefresh." + (reason ?? string.Empty));
            }
        }

        private static string FormatContextDescriptor(WorkbookContext context)
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

        private static string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return FormatWorkbookDescriptor(workbook, null);
        }

        private static string FormatWorkbookDescriptor(Excel.Workbook workbook, string fallbackFullName)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook, fallbackFullName)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private static string SafeWorkbookFullName(Excel.Workbook workbook, string fallbackFullName)
        {
            try
            {
                if (workbook == null)
                {
                    return fallbackFullName ?? string.Empty;
                }

                string fullName = workbook.FullName ?? string.Empty;
                return string.IsNullOrWhiteSpace(fullName) ? (fallbackFullName ?? string.Empty) : fullName;
            }
            catch
            {
                return fallbackFullName ?? string.Empty;
            }
        }

        private static string SafeWorkbookName(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null ? string.Empty : workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ResolveObservedWorkbookPath(WorkbookContext context, Excel.Workbook workbook)
        {
            if (context != null)
            {
                return SafeWorkbookFullName(context.Workbook, context.WorkbookFullName);
            }

            return SafeWorkbookFullName(workbook, null);
        }

        private static string FormatObservationCorrelationFields(WorkbookContext context, Excel.Workbook workbook)
        {
            Excel.Workbook observedWorkbook = context == null ? workbook : context.Workbook;
            return NewCaseVisibilityObservation.FormatCorrelationFields(
                null,
                observedWorkbook,
                ResolveObservedWorkbookPath(context, workbook));
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\"";
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.Hwnd.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
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
                return Convert.ToString(lateBoundWindow.Caption) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
