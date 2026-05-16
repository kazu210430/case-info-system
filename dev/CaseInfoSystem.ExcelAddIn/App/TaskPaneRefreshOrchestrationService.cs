using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshOrchestrationService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        internal const int PendingPaneRefreshIntervalMs = 400;
        internal const int PendingPaneRefreshMaxAttempts = 3;
        internal const int WorkbookPaneWindowResolveAttempts = 2;
        internal const int WorkbookPaneWindowResolveDelayMs = 80;

        private readonly ExcelInteropService _excelInteropService;
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly Logger _logger;
        private readonly TaskPaneRefreshCoordinator _taskPaneRefreshCoordinator;
        private readonly WorkbookTaskPaneReadyShowAttemptWorker _workbookTaskPaneReadyShowAttemptWorker;
        private readonly WorkbookPaneWindowResolver _workbookPaneWindowResolver;
        private readonly TaskPaneRefreshPreconditionDecisionService _preconditionDecisionService;
        private readonly TaskPaneRefreshObservationDecisionService _observationDecisionService;
        private readonly TaskPaneRefreshCompletionDecisionService _completionDecisionService;
        private readonly TaskPaneRefreshEmitPayloadBuilder _emitPayloadBuilder;
        private readonly TaskPaneForegroundGuaranteeTraceBuilder _foregroundTraceBuilder;
        private readonly TaskPaneRetryTimerLifecycle _retryTimerLifecycle;
        private readonly TaskPaneReadyShowRetryScheduler _readyShowRetryScheduler;
        private readonly WindowActivateDownstreamObservation _windowActivateDownstreamObservation;
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly ICasePaneHostBridge _casePaneHostBridge;
        private readonly PendingPaneRefreshRetryService _pendingPaneRefreshRetryService;
        private readonly object _createdCaseDisplaySessionSyncRoot = new object();
        private readonly Dictionary<string, CreatedCaseDisplaySession> _createdCaseDisplaySessions = new Dictionary<string, CreatedCaseDisplaySession>(StringComparer.OrdinalIgnoreCase);

        private int _kernelFlickerTraceRefreshAttemptSequence;
        private int _createdCaseDisplaySessionSequence;

        internal TaskPaneRefreshOrchestrationService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            TaskPaneRefreshCoordinator taskPaneRefreshCoordinator,
            WorkbookTaskPaneReadyShowAttemptWorker workbookTaskPaneReadyShowAttemptWorker,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount,
            ICasePaneHostBridge casePaneHostBridge)
        {
            _excelInteropService = excelInteropService;
            _workbookSessionService = workbookSessionService;
            _logger = logger;
            _taskPaneRefreshCoordinator = taskPaneRefreshCoordinator;
            _workbookTaskPaneReadyShowAttemptWorker = workbookTaskPaneReadyShowAttemptWorker ?? throw new ArgumentNullException(nameof(workbookTaskPaneReadyShowAttemptWorker));
            _workbookPaneWindowResolver = new WorkbookPaneWindowResolver(
                _excelInteropService,
                _logger,
                workbook => FormatWorkbookDescriptor(workbook),
                window => FormatWindowDescriptor(window),
                () => FormatActiveState());
            _preconditionDecisionService = new TaskPaneRefreshPreconditionDecisionService();
            _observationDecisionService = new TaskPaneRefreshObservationDecisionService();
            _completionDecisionService = new TaskPaneRefreshCompletionDecisionService();
            _emitPayloadBuilder = new TaskPaneRefreshEmitPayloadBuilder();
            _foregroundTraceBuilder = new TaskPaneForegroundGuaranteeTraceBuilder();
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
            _casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException(nameof(casePaneHostBridge));
            _retryTimerLifecycle = new TaskPaneRetryTimerLifecycle();
            _readyShowRetryScheduler = new TaskPaneReadyShowRetryScheduler(
                _logger,
                _retryTimerLifecycle,
                workbook => FormatWorkbookDescriptor(workbook),
                workbook => SafeWorkbookFullName(workbook));
            _windowActivateDownstreamObservation = new WindowActivateDownstreamObservation(
                _logger,
                workbook => FormatWorkbookDescriptor(workbook),
                window => FormatWindowDescriptor(window),
                () => FormatActiveState());
            _pendingPaneRefreshRetryService = new PendingPaneRefreshRetryService(
                _excelInteropService,
                _workbookSessionService,
                _logger,
                PendingPaneRefreshIntervalMs,
                PendingPaneRefreshMaxAttempts,
                _retryTimerLifecycle,
                TryRefreshTaskPane,
                ResolveWorkbookPaneWindow,
                StopPendingPaneRefreshTimer,
                workbook => FormatWorkbookDescriptor(workbook),
                window => FormatWindowDescriptor(window),
                () => FormatActiveState(),
                workbook => SafeWorkbookFullName(workbook),
                window => SafeWindowHwnd(window));
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return TryRefreshTaskPaneCore(reason, workbook, window, displayRequest: null);
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(TaskPaneDisplayRequest displayRequest, Excel.Workbook workbook, Excel.Window window)
        {
            string reason = displayRequest == null ? string.Empty : displayRequest.ToReasonString();
            return TryRefreshTaskPaneCore(reason, workbook, window, displayRequest);
        }

        private TaskPaneRefreshAttemptResult TryRefreshTaskPaneCore(string reason, Excel.Workbook workbook, Excel.Window window, TaskPaneDisplayRequest displayRequest)
        {
            TaskPaneRefreshAttemptStartObservation attemptObservation = StartTaskPaneRefreshAttemptObservation(
                reason,
                workbook,
                window,
                displayRequest);
            TaskPaneRefreshPreconditionDecisionResult preconditionDecision = EvaluateTaskPaneRefreshPreconditionBoundary(
                reason,
                workbook,
                window,
                attemptObservation.RefreshAttemptId);
            if (!preconditionDecision.CanRefresh)
            {
                return ReturnFailClosedTaskPaneRefreshResult(
                    reason,
                    workbook,
                    window,
                    displayRequest,
                    preconditionDecision,
                    attemptObservation);
            }

            RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(
                reason,
                workbook,
                window);
            return ContinuePostDispatchRefreshConvergence(
                reason,
                workbook,
                window,
                displayRequest,
                routeDispatchExecutionResult,
                attemptObservation.Stopwatch,
                attemptObservation.RefreshAttemptId);
        }

        private TaskPaneRefreshAttemptStartObservation StartTaskPaneRefreshAttemptObservation(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            TaskPaneDisplayRequest displayRequest)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            int refreshAttemptId = ++_kernelFlickerTraceRefreshAttemptSequence;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=try-refresh-start refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", activeState="
                + FormatActiveState()
                + WindowActivateDownstreamObservation.FormatDisplayRequestTraceFields(displayRequest));
            _windowActivateDownstreamObservation.LogStart(displayRequest, reason, workbook, window, refreshAttemptId);
            return new TaskPaneRefreshAttemptStartObservation(stopwatch, refreshAttemptId);
        }

        private readonly struct TaskPaneRefreshAttemptStartObservation
        {
            internal TaskPaneRefreshAttemptStartObservation(Stopwatch stopwatch, int refreshAttemptId)
            {
                Stopwatch = stopwatch;
                RefreshAttemptId = refreshAttemptId;
            }

            internal Stopwatch Stopwatch { get; }

            internal int RefreshAttemptId { get; }
        }

        private TaskPaneRefreshPreconditionDecisionResult EvaluateTaskPaneRefreshPreconditionBoundary(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            int refreshAttemptId)
        {
            TaskPaneRefreshPreconditionDecisionResult preconditionDecision = _preconditionDecisionService.Decide(
                reason,
                workbook,
                window,
                () => _casePaneHostBridge.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window));
            if (!preconditionDecision.CanRefresh)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action="
                    + preconditionDecision.SkipActionName
                    + " refreshAttemptId="
                    + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", inputWindow="
                    + FormatWindowDescriptor(window)
                    + ", activeState="
                    + FormatActiveState());
            }

            return preconditionDecision;
        }

        private TaskPaneRefreshAttemptResult ReturnFailClosedTaskPaneRefreshResult(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            TaskPaneDisplayRequest displayRequest,
            TaskPaneRefreshPreconditionDecisionResult preconditionDecision,
            TaskPaneRefreshAttemptStartObservation attemptObservation)
        {
            TaskPaneRefreshAttemptResult skippedResult = CompleteNormalizedOutcomeChain(
                reason,
                workbook,
                window,
                preconditionDecision.NormalizedOutcomeAttemptResult,
                attemptObservation.Stopwatch,
                preconditionDecision.NormalizedOutcomeActionName,
                null,
                null);
            _windowActivateDownstreamObservation.LogOutcome(
                displayRequest,
                reason,
                skippedResult,
                attemptObservation.Stopwatch,
                attemptObservation.RefreshAttemptId,
                preconditionDecision.SkipActionName);
            return skippedResult;
        }

        private RefreshDispatchExecutionResult DispatchTaskPaneRefreshRoute(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            RefreshDispatchExecutionResult dispatchExecutionResult = RefreshDispatchShell.Dispatch(
                _taskPaneRefreshCoordinator,
                reason,
                workbook,
                window,
                _getKernelHomeForm,
                _getTaskPaneRefreshSuppressionCount);
            return dispatchExecutionResult;
        }

        private TaskPaneRefreshAttemptResult ContinuePostDispatchRefreshConvergence(
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            TaskPaneDisplayRequest displayRequest,
            RefreshDispatchExecutionResult routeDispatchExecutionResult,
            Stopwatch stopwatch,
            int refreshAttemptId)
        {
            TaskPaneRefreshAttemptResult attemptResult = CompleteNormalizedOutcomeChain(
                reason,
                workbook,
                window,
                routeDispatchExecutionResult.AttemptResult,
                stopwatch,
                "refresh",
                null,
                null);
            attemptResult = CompleteForegroundGuaranteeOutcome(
                reason,
                workbook,
                window,
                attemptResult,
                stopwatch);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=try-refresh-end refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", result="
                + RefreshDispatchExecutionResult.FormatResultText(attemptResult));
            _windowActivateDownstreamObservation.LogOutcome(
                displayRequest,
                reason,
                attemptResult,
                stopwatch,
                refreshAttemptId,
                "refresh");
            TryCompleteCreatedCaseDisplaySession(
                null,
                reason,
                workbook,
                window,
                attemptResult,
                "refresh",
                displayRequest: displayRequest);
            return attemptResult;
        }

        internal bool IsTaskPaneRefreshSucceeded(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return TryRefreshTaskPane(reason, workbook, window).IsRefreshSucceeded;
        }

        internal void RefreshActiveTaskPane(string reason)
        {
            TryRefreshTaskPane(reason, null, null);
        }

        // Handoff flow: entries, shared branch routing, branch helpers, then data-only inputs.
        internal void ScheduleActiveTaskPaneRefresh(string reason)
        {
            RunTaskPaneRefreshHandoffFlow(TaskPaneRefreshHandoffFlowInput.ForActiveRefresh(reason));
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
        {
            RunTaskPaneRefreshHandoffFlow(TaskPaneRefreshHandoffFlowInput.ForWorkbookFallback(workbook, reason));
        }

        private void RunTaskPaneRefreshHandoffFlow(TaskPaneRefreshHandoffFlowInput flowInput)
        {
            if (flowInput.IsActiveRefresh)
            {
                RunActiveTaskPaneRefreshHandoffBranch(flowInput);
                return;
            }

            RunWorkbookFallbackTaskPaneRefreshHandoffBranch(flowInput);
        }

        private void RunActiveTaskPaneRefreshHandoffBranch(TaskPaneRefreshHandoffFlowInput flowInput)
        {
            ActiveTaskPaneRefreshHandoff activeHandoff = BeginActiveTaskPaneRefreshHandoff(flowInput.Reason);
            if (TryRefreshActiveTaskPaneImmediately(activeHandoff))
            {
                return;
            }

            StartPendingRefreshRetryFromActiveHandoff(activeHandoff);
        }

        private void RunWorkbookFallbackTaskPaneRefreshHandoffBranch(TaskPaneRefreshHandoffFlowInput flowInput)
        {
            PendingFallbackRefreshHandoff fallbackHandoff = BeginPendingFallbackRefreshHandoff(flowInput.WorkbookFallbackWorkbook, flowInput.Reason);
            if (ShouldSkipPendingFallbackForWorkbookOpenBoundary(fallbackHandoff))
            {
                LogPendingFallbackWorkbookOpenSkip(fallbackHandoff);
                return;
            }

            PendingRefreshRetryHandoff retryHandoff = PreparePendingRefreshRetryHandoff(fallbackHandoff);
            if (TryRefreshPendingFallbackImmediately(retryHandoff))
            {
                return;
            }

            StartPendingRefreshRetryFromFallback(retryHandoff);
        }

        private ActiveTaskPaneRefreshHandoff BeginActiveTaskPaneRefreshHandoff(string reason)
        {
            _pendingPaneRefreshRetryService.TrackActiveTarget();
            return new ActiveTaskPaneRefreshHandoff(reason);
        }

        private bool TryRefreshActiveTaskPaneImmediately(ActiveTaskPaneRefreshHandoff activeHandoff)
        {
            if (!IsTaskPaneRefreshSucceeded(activeHandoff.Reason, null, null))
            {
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-immediate-success reason="
                + (activeHandoff.Reason ?? string.Empty)
                + ", target=active");
            StopPendingPaneRefreshTimer();
            return true;
        }

        private void StartPendingRefreshRetryFromActiveHandoff(ActiveTaskPaneRefreshHandoff activeHandoff)
        {
            int attemptsRemaining = _pendingPaneRefreshRetryService.BeginRetrySequence(activeHandoff.Reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (activeHandoff.Reason ?? string.Empty)
                + ", target=active"
                + ", attempts="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        private PendingFallbackRefreshHandoff BeginPendingFallbackRefreshHandoff(Excel.Workbook workbook, string reason)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=wait-ready-fallback-handoff reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", maxAttempts="
                + WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true"
                + ", activeState="
                + FormatActiveState()
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            _logger?.Info(
                "TaskPane wait-ready fallback handoff. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", maxAttempts="
                + WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts.ToString(CultureInfo.InvariantCulture)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true"
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                null,
                "ready-show-fallback-handoff",
                "TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh",
                SafeWorkbookFullName(workbook),
                "reason=" + (reason ?? string.Empty) + ",fallbackCause=AttemptsExhausted");

            return new PendingFallbackRefreshHandoff(workbook, reason);
        }

        private static bool ShouldSkipPendingFallbackForWorkbookOpenBoundary(PendingFallbackRefreshHandoff fallbackHandoff)
        {
            return TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(
                fallbackHandoff.Reason,
                fallbackHandoff.Workbook,
                window: null);
        }

        private void LogPendingFallbackWorkbookOpenSkip(PendingFallbackRefreshHandoff fallbackHandoff)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=skip-workbook-open-defer reason="
                + (fallbackHandoff.Reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(fallbackHandoff.Workbook)
                + ", activeState="
                + FormatActiveState());
        }

        private PendingRefreshRetryHandoff PreparePendingRefreshRetryHandoff(PendingFallbackRefreshHandoff fallbackHandoff)
        {
            Excel.Workbook workbook = fallbackHandoff.Workbook;
            string reason = fallbackHandoff.Reason;
            _pendingPaneRefreshRetryService.TrackWorkbookTarget(_excelInteropService == null
                ? string.Empty
                : _excelInteropService.GetWorkbookFullName(workbook));
            Excel.Window workbookWindow = ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook: false);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-prepare reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", windowResolved="
                + (workbookWindow != null).ToString()
                + ", resolvedWindow="
                + FormatWindowDescriptor(workbookWindow)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true"
                + ", activeState="
                + FormatActiveState());
            _logger?.Info(
                "TaskPane timer fallback prepare. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", windowResolved="
                + (workbookWindow != null).ToString()
                + ", resolvedWindowHwnd="
                + SafeWindowHwnd(workbookWindow)
                + ", fallbackCause=AttemptsExhausted"
                + ", fallbackHandoff=true");

            return new PendingRefreshRetryHandoff(workbook, reason, workbookWindow);
        }

        private bool TryRefreshPendingFallbackImmediately(PendingRefreshRetryHandoff retryHandoff)
        {
            if (!TryRefreshTaskPane(retryHandoff.Reason, retryHandoff.Workbook, retryHandoff.WorkbookWindow).IsRefreshSucceeded)
            {
                return false;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-immediate-success reason="
                + (retryHandoff.Reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(retryHandoff.Workbook));
            _logger?.Info("TaskPane timer fallback immediate refresh succeeded. reason=" + (retryHandoff.Reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(retryHandoff.Workbook));
            StopPendingPaneRefreshTimer();
            return true;
        }

        private void StartPendingRefreshRetryFromFallback(PendingRefreshRetryHandoff retryHandoff)
        {
            int attemptsRemaining = _pendingPaneRefreshRetryService.BeginRetrySequence(retryHandoff.Reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (retryHandoff.Reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(retryHandoff.Workbook)
                + ", attempts="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
            _logger?.Info("TaskPane timer fallback scheduled. reason=" + (retryHandoff.Reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(retryHandoff.Workbook) + ", attempts=" + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        private readonly struct TaskPaneRefreshHandoffFlowInput
        {
            private TaskPaneRefreshHandoffFlowInput(bool isActiveRefresh, Excel.Workbook workbookFallbackWorkbook, string reason)
            {
                IsActiveRefresh = isActiveRefresh;
                WorkbookFallbackWorkbook = workbookFallbackWorkbook;
                Reason = reason;
            }

            internal bool IsActiveRefresh { get; }

            internal Excel.Workbook WorkbookFallbackWorkbook { get; }

            internal string Reason { get; }

            internal static TaskPaneRefreshHandoffFlowInput ForActiveRefresh(string reason)
            {
                return new TaskPaneRefreshHandoffFlowInput(isActiveRefresh: true, workbookFallbackWorkbook: null, reason: reason);
            }

            internal static TaskPaneRefreshHandoffFlowInput ForWorkbookFallback(Excel.Workbook workbook, string reason)
            {
                return new TaskPaneRefreshHandoffFlowInput(isActiveRefresh: false, workbookFallbackWorkbook: workbook, reason: reason);
            }
        }

        private readonly struct ActiveTaskPaneRefreshHandoff
        {
            internal ActiveTaskPaneRefreshHandoff(string reason)
            {
                Reason = reason;
            }

            internal string Reason { get; }
        }

        private readonly struct PendingFallbackRefreshHandoff
        {
            internal PendingFallbackRefreshHandoff(Excel.Workbook workbook, string reason)
            {
                Workbook = workbook;
                Reason = reason;
            }

            internal Excel.Workbook Workbook { get; }

            internal string Reason { get; }
        }

        private readonly struct PendingRefreshRetryHandoff
        {
            internal PendingRefreshRetryHandoff(Excel.Workbook workbook, string reason, Excel.Window workbookWindow)
            {
                Workbook = workbook;
                Reason = reason;
                WorkbookWindow = workbookWindow;
            }

            internal Excel.Workbook Workbook { get; }

            internal string Reason { get; }

            internal Excel.Window WorkbookWindow { get; }
        }

        internal void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=wait-ready-enqueued reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveState()
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            _logger?.Info(
                "TaskPane wait-ready enqueued. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                null,
                "ready-show-enqueued",
                "TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady",
                SafeWorkbookFullName(workbook),
                "reason=" + (reason ?? string.Empty));
            CreatedCaseDisplaySession createdCaseDisplaySession = BeginCreatedCaseDisplaySession(workbook, reason);
            _workbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(
                workbook,
                reason,
                _readyShowRetryScheduler.Schedule,
                outcome => HandleWorkbookTaskPaneShown(createdCaseDisplaySession, workbook, reason, outcome),
                ScheduleWorkbookTaskPaneRefresh);
        }

        internal Excel.Window ResolveWorkbookPaneWindow(Excel.Workbook workbook, string reason, bool activateWorkbook)
        {
            return _workbookPaneWindowResolver.Resolve(workbook, reason, activateWorkbook);
        }

        internal void StopPendingPaneRefreshTimer()
        {
            _pendingPaneRefreshRetryService.StopTimer();
            _retryTimerLifecycle.StopWaitReadyRetryTimers();
        }

        private TaskPaneRefreshAttemptResult CompleteNormalizedOutcomeChain(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            TaskPaneRefreshObservationDecision decision = _observationDecisionService.CompleteNormalizedOutcomeChain(
                new TaskPaneRefreshObservationDecisionInput(
                    reason,
                    workbook,
                    inputWindow,
                    attemptResult,
                    completionSource,
                    attemptNumber,
                    workbookWindowEnsureFacts));
            LogVisibilityRecoveryOutcome(
                reason,
                workbook,
                inputWindow,
                decision.Visibility,
                stopwatch);
            LogRefreshSourceSelectionOutcome(
                reason,
                workbook,
                decision.RefreshSource,
                stopwatch,
                inputWindow);
            LogRebuildFallbackOutcome(
                reason,
                workbook,
                decision.RebuildFallback,
                stopwatch,
                inputWindow);
            return decision.AttemptResult;
        }

        private void LogVisibilityRecoveryOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshVisibilityObservationDecision decision,
            Stopwatch stopwatch)
        {
            if (!IsCreatedCaseDisplayReason(reason) || decision == null || decision.Outcome == null)
            {
                return;
            }

            VisibilityRecoveryOutcome outcome = decision.Outcome;
            WorkbookContext context = decision.Context;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=visibility-recovery-decision reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", inputWindow="
                + FormatWindowDescriptor(inputWindow)
                + ", visibilityRecoveryStatus="
                + outcome.Status.ToString()
                + ", visibilityRecoveryReason="
                + outcome.Reason
                + ", visibilityRecoveryDisplayCompletable="
                + outcome.IsDisplayCompletable.ToString()
                + ", paneVisible="
                + outcome.IsPaneVisible.ToString()
                + ", paneVisibleSource="
                + outcome.PaneVisibleSource.ToString()
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + FormatObservationCorrelationFields(context, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                decision.ObservedWindow,
                "visibility-recovery-decision",
                "TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                decision.Details);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                decision.ObservedWindow,
                "visibility-recovery-" + outcome.Status.ToString().ToLowerInvariant(),
                "TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                decision.Details);
        }

        private void LogRefreshSourceSelectionOutcome(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshSourceObservationDecision decision,
            Stopwatch stopwatch,
            Excel.Window inputWindow)
        {
            if (!IsCreatedCaseDisplayReason(reason) || decision == null || decision.Outcome == null)
            {
                return;
            }

            RefreshSourceSelectionOutcome outcome = decision.Outcome;
            LogRefreshSourceSelectionTrace(
                reason,
                workbook,
                decision.Context,
                decision.ObservedWindow,
                outcome,
                stopwatch,
                decision.StatusAction,
                decision.Details);

            if (decision.ShouldLogRebuildRequiredTrace)
            {
                LogRefreshSourceSelectionTrace(
                    reason,
                    workbook,
                    decision.Context,
                    decision.ObservedWindow,
                    outcome,
                    stopwatch,
                    "refresh-source-rebuild-required",
                    decision.Details);
            }
        }

        private void LogRefreshSourceSelectionTrace(
            string reason,
            Excel.Workbook workbook,
            WorkbookContext context,
            Excel.Window observedWindow,
            RefreshSourceSelectionOutcome outcome,
            Stopwatch stopwatch,
            string statusAction,
            string detail)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action="
                + statusAction
                + " reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", refreshSourceStatus="
                + outcome.Status.ToString()
                + ", selectedSource="
                + outcome.SelectedSource.ToString()
                + ", selectionReason="
                + outcome.SelectionReason
                + ", cacheFallback="
                + outcome.IsCacheFallback.ToString()
                + ", rebuildRequired="
                + outcome.IsRebuildRequired.ToString()
                + ", canContinue="
                + outcome.CanContinueRefresh.ToString()
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + FormatObservationCorrelationFields(context, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                observedWindow,
                statusAction,
                "TaskPaneRefreshOrchestrationService.CompleteRefreshSourceSelectionOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                detail);
        }

        private void LogRebuildFallbackOutcome(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshRebuildFallbackObservationDecision decision,
            Stopwatch stopwatch,
            Excel.Window inputWindow)
        {
            if (!IsCreatedCaseDisplayReason(reason) || decision == null || decision.Outcome == null)
            {
                return;
            }

            RebuildFallbackOutcome outcome = decision.Outcome;
            WorkbookContext context = decision.Context;
            if (outcome.IsRequired)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=rebuild-fallback-required reason="
                    + (reason ?? string.Empty)
                    + ", context="
                    + FormatContextDescriptor(context)
                    + ", snapshotSource="
                    + outcome.SnapshotSource.ToString()
                    + ", fallbackReasons="
                    + outcome.FallbackReasons
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                    + FormatObservationCorrelationFields(context, workbook));
                NewCaseVisibilityObservation.Log(
                    _logger,
                    _excelInteropService,
                    null,
                    context == null ? workbook : context.Workbook,
                    decision.ObservedWindow,
                    "rebuild-fallback-required",
                    "TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome",
                    ResolveObservedWorkbookPath(context, workbook),
                    decision.Details);
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action="
                + decision.StatusAction
                + " reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", rebuildFallbackStatus="
                + outcome.Status.ToString()
                + ", rebuildFallbackRequired="
                + outcome.IsRequired.ToString()
                + ", rebuildFallbackCanContinue="
                + outcome.CanContinueRefresh.ToString()
                + ", snapshotSource="
                + outcome.SnapshotSource.ToString()
                + ", fallbackReasons="
                + outcome.FallbackReasons
                + ", failureReason="
                + outcome.FailureReason
                + ", degradedReason="
                + outcome.DegradedReason
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + FormatObservationCorrelationFields(context, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                decision.ObservedWindow,
                decision.StatusAction,
                "TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                decision.Details);
        }

        private TaskPaneRefreshAttemptResult CompleteForegroundGuaranteeOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch)
        {
            if (attemptResult == null)
            {
                return null;
            }

            TaskPaneRefreshForegroundGuaranteeDecision foregroundDecision =
                _observationDecisionService.DecideForegroundGuarantee(attemptResult, inputWindow);
            LogForegroundGuaranteeDecision(reason, workbook, foregroundDecision, stopwatch);
            if (!foregroundDecision.ShouldExecuteForegroundGuarantee)
            {
                return foregroundDecision.AttemptResult;
            }

            ForegroundGuaranteeOutcome requiredOutcome = ExecuteForegroundGuaranteeAndBuildOutcome(
                reason,
                workbook,
                foregroundDecision,
                stopwatch);
            return attemptResult.WithForegroundGuaranteeOutcome(requiredOutcome);
        }

        private ForegroundGuaranteeOutcome ExecuteForegroundGuaranteeAndBuildOutcome(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshForegroundGuaranteeDecision foregroundDecision,
            Stopwatch stopwatch)
        {
            TaskPaneRefreshAttemptResult attemptResult = foregroundDecision.AttemptResult;
            LogFinalForegroundGuaranteeStarted(reason, workbook, attemptResult, stopwatch);
            ForegroundGuaranteeExecutionResult executionResult = _taskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(
                attemptResult.ForegroundContext,
                attemptResult.ForegroundWorkbook,
                reason);
            NewCaseDefaultTimingLogHelper.LogDetail(
                _logger,
                ResolveObservedWorkbookPath(attemptResult.ForegroundContext, attemptResult.ForegroundWorkbook),
                "waitUiCloseToFinalForegroundStable",
                "tryRecoverWorkbookWindow",
                executionResult.ElapsedMilliseconds,
                "reason=" + (reason ?? string.Empty));
            LogFinalForegroundGuaranteeCompleted(reason, workbook, attemptResult, executionResult, stopwatch);
            _taskPaneRefreshCoordinator.BeginPostForegroundProtection(
                attemptResult.ForegroundContext,
                attemptResult.ForegroundWorkbook,
                reason,
                stopwatch.ElapsedMilliseconds);
            NewCaseDefaultTimingLogHelper.LogWaitUiCloseToFinalForegroundStable(
                _logger,
                ResolveObservedWorkbookPath(attemptResult.ForegroundContext, attemptResult.ForegroundWorkbook),
                reason,
                executionResult.Recovered);

            return _observationDecisionService.ClassifyRequiredForegroundExecutionOutcome(
                foregroundDecision.TargetKind,
                executionResult);
        }

        private void LogForegroundGuaranteeDecision(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshForegroundGuaranteeDecision decision,
            Stopwatch stopwatch)
        {
            WorkbookContext context = decision == null ? null : decision.Context;
            TaskPaneForegroundGuaranteeTracePayload trace = _foregroundTraceBuilder.BuildDecisionTrace(
                new TaskPaneForegroundGuaranteeDecisionTraceInput(
                    reason,
                    decision,
                    FormatContextDescriptor(context),
                    stopwatch.ElapsedMilliseconds,
                    FormatObservationCorrelationFields(context, workbook)));
            _logger?.Info(trace.KernelTraceMessage);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                decision == null ? null : decision.ObservedWindow,
                trace.ObservationAction,
                trace.ObservationSource,
                ResolveObservedWorkbookPath(context, workbook),
                trace.Details);
        }

        private void LogFinalForegroundGuaranteeStarted(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            TaskPaneForegroundGuaranteeTracePayload trace = _foregroundTraceBuilder.BuildStartedTrace(
                new TaskPaneForegroundGuaranteeStartedTraceInput(
                    reason,
                    FormatContextDescriptor(context),
                    stopwatch.ElapsedMilliseconds,
                    FormatObservationCorrelationFields(context, workbook)));
            _logger?.Info(trace.KernelTraceMessage);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                context == null ? null : context.Window,
                trace.ObservationAction,
                trace.ObservationSource,
                ResolveObservedWorkbookPath(context, workbook),
                trace.Details);
        }

        private void LogFinalForegroundGuaranteeCompleted(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeExecutionResult executionResult,
            Stopwatch stopwatch)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            TaskPaneForegroundGuaranteeTracePayload trace = _foregroundTraceBuilder.BuildCompletedTrace(
                new TaskPaneForegroundGuaranteeCompletedTraceInput(
                    reason,
                    executionResult,
                    FormatContextDescriptor(context),
                    stopwatch.ElapsedMilliseconds,
                    FormatObservationCorrelationFields(context, workbook)));
            _logger?.Info(trace.KernelTraceMessage);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                context == null ? null : context.Window,
                trace.ObservationAction,
                trace.ObservationSource,
                ResolveObservedWorkbookPath(context, workbook),
                trace.Details);
        }

        private CreatedCaseDisplaySession BeginCreatedCaseDisplaySession(Excel.Workbook workbook, string reason)
        {
            if (!IsCreatedCaseDisplayReason(reason) || workbook == null)
            {
                return null;
            }

            string workbookFullName = SafeWorkbookFullName(workbook);
            if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                return null;
            }

            CreatedCaseDisplaySession session = new CreatedCaseDisplaySession(
                "CDS-" + (++_createdCaseDisplaySessionSequence).ToString("0000", CultureInfo.InvariantCulture),
                workbookFullName,
                reason);
            lock (_createdCaseDisplaySessionSyncRoot)
            {
                _createdCaseDisplaySessions[workbookFullName] = session;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=created-case-display-session-started sessionId="
                + session.SessionId
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                null,
                "created-case-display-session-started",
                "TaskPaneRefreshOrchestrationService.BeginCreatedCaseDisplaySession",
                workbookFullName,
                "reason=" + (reason ?? string.Empty) + ",sessionId=" + session.SessionId);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                null,
                "display-handoff-completed",
                "TaskPaneRefreshOrchestrationService.BeginCreatedCaseDisplaySession",
                workbookFullName,
                "reason=" + (reason ?? string.Empty) + ",sessionId=" + session.SessionId);
            return session;
        }

        private void HandleWorkbookTaskPaneShown(
            CreatedCaseDisplaySession session,
            Excel.Workbook workbook,
            string reason,
            WorkbookTaskPaneReadyShowAttemptOutcome outcome)
        {
            StopPendingPaneRefreshTimer();
            if (outcome == null)
            {
                return;
            }

            ReadyShowCallbackFacts callbackFacts = BuildReadyShowCallbackFacts(outcome);
            Stopwatch stopwatch = Stopwatch.StartNew();
            TaskPaneRefreshAttemptResult attemptResult = CompleteNormalizedOutcomeChain(
                reason,
                workbook,
                callbackFacts.WorkbookWindow,
                callbackFacts.RefreshAttemptResult,
                stopwatch,
                "ready-show-attempt",
                callbackFacts.AttemptNumber,
                callbackFacts.WorkbookWindowEnsureFacts);
            attemptResult = CompleteForegroundGuaranteeOutcome(
                reason,
                workbook,
                callbackFacts.WorkbookWindow,
                attemptResult,
                stopwatch);
            TryCompleteCreatedCaseDisplaySession(
                session,
                reason,
                workbook,
                callbackFacts.WorkbookWindow,
                attemptResult,
                "ready-show-attempt",
                callbackFacts.AttemptNumber);
        }

        private void TryCompleteCreatedCaseDisplaySession(
            CreatedCaseDisplaySession session,
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber = null,
            TaskPaneDisplayRequest displayRequest = null)
        {
            CreatedCaseDisplayCompletionDecision completionDecision =
                _completionDecisionService.DecideCreatedCaseDisplayCompletion(
                    new TaskPaneRefreshCompletionDecisionInput(
                        IsCreatedCaseDisplayReason(reason),
                        attemptResult));
            if (!completionDecision.CanComplete)
            {
                return;
            }

            CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);
            if (resolvedSession == null)
            {
                return;
            }

            if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))
            {
                return;
            }

            CaseDisplayCompletedPayload payload = _emitPayloadBuilder.BuildCaseDisplayCompleted(
                new CaseDisplayCompletedPayloadInput(
                    reason,
                    resolvedSession.SessionId,
                    resolvedSession.WorkbookFullName,
                    attemptResult,
                    completionSource,
                    attemptNumber,
                    displayRequest,
                    FormatWorkbookDescriptor(workbook),
                    FormatWindowDescriptor(window)));

            _logger?.Info(payload.KernelTraceMessage);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                window,
                payload.ObservationAction,
                payload.ObservationSource,
                payload.WorkbookFullName,
                payload.Details);
            NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);
        }

        private bool TryMarkCreatedCaseDisplaySessionCompletedForEmit(CreatedCaseDisplaySession resolvedSession)
        {
            bool shouldEmit = false;
            lock (_createdCaseDisplaySessionSyncRoot)
            {
                if (!resolvedSession.IsCompleted)
                {
                    resolvedSession.IsCompleted = true;
                    _createdCaseDisplaySessions.Remove(resolvedSession.WorkbookFullName);
                    shouldEmit = true;
                }
            }

            return shouldEmit;
        }

        private CreatedCaseDisplaySession ResolveCreatedCaseDisplaySession(string reason, Excel.Workbook workbook)
        {
            if (!IsCreatedCaseDisplayReason(reason))
            {
                return null;
            }

            string workbookFullName = SafeWorkbookFullName(workbook);
            lock (_createdCaseDisplaySessionSyncRoot)
            {
                if (!string.IsNullOrWhiteSpace(workbookFullName)
                    && _createdCaseDisplaySessions.TryGetValue(workbookFullName, out CreatedCaseDisplaySession session))
                {
                    return session;
                }

                if (_createdCaseDisplaySessions.Count == 1)
                {
                    foreach (CreatedCaseDisplaySession activeSession in _createdCaseDisplaySessions.Values)
                    {
                        return activeSession;
                    }
                }
            }

            return null;
        }

        private static bool IsCreatedCaseDisplayReason(string reason)
        {
            return string.Equals(reason, ControlFlowReasons.CreatedCasePostRelease, StringComparison.OrdinalIgnoreCase);
        }

        private static ReadyShowCallbackFacts BuildReadyShowCallbackFacts(WorkbookTaskPaneReadyShowAttemptOutcome outcome)
        {
            return new ReadyShowCallbackFacts(
                outcome.WorkbookWindow,
                outcome.RefreshAttemptResult,
                outcome.AttemptNumber,
                outcome.WorkbookWindowEnsureFacts);
        }

        private readonly struct ReadyShowCallbackFacts
        {
            internal ReadyShowCallbackFacts(
                Excel.Window workbookWindow,
                TaskPaneRefreshAttemptResult refreshAttemptResult,
                int attemptNumber,
                WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
            {
                WorkbookWindow = workbookWindow;
                RefreshAttemptResult = refreshAttemptResult;
                AttemptNumber = attemptNumber;
                WorkbookWindowEnsureFacts = workbookWindowEnsureFacts;
            }

            internal Excel.Window WorkbookWindow { get; }

            internal TaskPaneRefreshAttemptResult RefreshAttemptResult { get; }

            internal int AttemptNumber { get; }

            internal WorkbookWindowVisibilityEnsureFacts WorkbookWindowEnsureFacts { get; }
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
                + FormatWorkbookDescriptor(context.Workbook)
                + ",window="
                + FormatWindowDescriptor(context.Window)
                + ",activeSheet=\""
                + (context.ActiveSheetCodeName ?? string.Empty)
                + "\"";
        }

        private string ResolveObservedWorkbookPath(WorkbookContext context, Excel.Workbook workbook)
        {
            if (context != null && !string.IsNullOrWhiteSpace(context.WorkbookFullName))
            {
                return context.WorkbookFullName;
            }

            return SafeWorkbookFullName(context == null ? workbook : context.Workbook);
        }

        private string FormatObservationCorrelationFields(WorkbookContext context, Excel.Workbook workbook)
        {
            Excel.Workbook observedWorkbook = context == null ? workbook : context.Workbook;
            return NewCaseVisibilityObservation.FormatCorrelationFields(
                _excelInteropService,
                observedWorkbook,
                ResolveObservedWorkbookPath(context, workbook));
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        private sealed class CreatedCaseDisplaySession
        {
            internal CreatedCaseDisplaySession(string sessionId, string workbookFullName, string reason)
            {
                SessionId = sessionId ?? string.Empty;
                WorkbookFullName = workbookFullName ?? string.Empty;
                Reason = reason ?? string.Empty;
            }

            internal string SessionId { get; }

            internal string WorkbookFullName { get; }

            internal string Reason { get; }

            internal bool IsCompleted { get; set; }
        }

        private static class RefreshDispatchShell
        {
            internal static RefreshDispatchExecutionResult Dispatch(
                TaskPaneRefreshCoordinator taskPaneRefreshCoordinator,
                string reason,
                Excel.Workbook workbook,
                Excel.Window window,
                Func<KernelHomeForm> getKernelHomeForm,
                Func<int> getTaskPaneRefreshSuppressionCount)
            {
                TaskPaneRefreshAttemptResult attemptResult = taskPaneRefreshCoordinator.TryRefreshTaskPane(
                    reason,
                    workbook,
                    window,
                    getKernelHomeForm == null ? null : getKernelHomeForm(),
                    getTaskPaneRefreshSuppressionCount == null ? 0 : getTaskPaneRefreshSuppressionCount());
                return RefreshDispatchExecutionResult.FromAttemptResult(attemptResult);
            }
        }

        private sealed class RefreshDispatchExecutionResult
        {
            private RefreshDispatchExecutionResult(TaskPaneRefreshAttemptResult attemptResult, string resultText)
            {
                AttemptResult = attemptResult;
                ResultText = resultText ?? string.Empty;
            }

            internal TaskPaneRefreshAttemptResult AttemptResult { get; }

            internal string ResultText { get; }

            internal static RefreshDispatchExecutionResult FromAttemptResult(TaskPaneRefreshAttemptResult attemptResult)
            {
                return new RefreshDispatchExecutionResult(
                    attemptResult,
                    FormatResultText(attemptResult));
            }

            internal static string FormatResultText(TaskPaneRefreshAttemptResult attemptResult)
            {
                if (attemptResult == null)
                {
                    return "null";
                }

                ForegroundGuaranteeOutcome outcome = attemptResult.ForegroundGuaranteeOutcome;
                return attemptResult.IsRefreshSucceeded.ToString()
                    + ",foregroundOutcome="
                    + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString());
            }
        }

        private sealed class WorkbookPaneWindowResolver
        {
            private readonly ExcelInteropService _excelInteropService;
            private readonly Logger _logger;
            private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
            private readonly Func<Excel.Window, string> _formatWindowDescriptor;
            private readonly Func<string> _formatActiveState;

            internal WorkbookPaneWindowResolver(
                ExcelInteropService excelInteropService,
                Logger logger,
                Func<Excel.Workbook, string> formatWorkbookDescriptor,
                Func<Excel.Window, string> formatWindowDescriptor,
                Func<string> formatActiveState)
            {
                _excelInteropService = excelInteropService;
                _logger = logger;
                _formatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
                _formatWindowDescriptor = formatWindowDescriptor ?? throw new ArgumentNullException(nameof(formatWindowDescriptor));
                _formatActiveState = formatActiveState ?? throw new ArgumentNullException(nameof(formatActiveState));
            }

            internal Excel.Window Resolve(Excel.Workbook workbook, string reason, bool activateWorkbook)
            {
                if (_excelInteropService == null || workbook == null)
                {
                    return null;
                }

                string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=resolve-window-start reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook)
                    + ", activateWorkbook="
                    + activateWorkbook.ToString()
                    + ", activeState="
                    + _formatActiveState());
                for (int attempt = 0; attempt < WorkbookPaneWindowResolveAttempts; attempt++)
                {
                    if (activateWorkbook)
                    {
                        _excelInteropService.ActivateWorkbook(workbook);
                    }

                    Excel.Window workbookWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
                    Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();
                    string activeWorkbookFullName = _excelInteropService.GetWorkbookFullName(activeWorkbook);
                    Excel.Window activeWindow = _excelInteropService.GetActiveWindow();
                    bool activeWorkbookMatches = string.Equals(activeWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase);

                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneRefreshOrchestrationService action=resolve-window-state reason="
                        + (reason ?? string.Empty)
                        + ", workbookFullName=\""
                        + workbookFullName
                        + "\", resolveAttempt="
                        + (attempt + 1).ToString(CultureInfo.InvariantCulture)
                        + ", activateWorkbook="
                        + activateWorkbook.ToString()
                        + ", visibleWindow="
                        + _formatWindowDescriptor(workbookWindow)
                        + ", activeWorkbook="
                        + _formatWorkbookDescriptor(activeWorkbook)
                        + ", activeWindow="
                        + _formatWindowDescriptor(activeWindow)
                        + ", activeWorkbookMatches="
                        + activeWorkbookMatches.ToString());
                    _logger?.Info("ResolveWorkbookPaneWindow state. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", resolveAttempt=" + (attempt + 1).ToString(CultureInfo.InvariantCulture) + ", activateWorkbook=" + activateWorkbook.ToString() + ", visibleWindowHwnd=" + SafeWindowHwnd(workbookWindow) + ", activeWorkbook=" + activeWorkbookFullName + ", activeWorkbookMatches=" + activeWorkbookMatches.ToString() + ", activeWindowHwnd=" + SafeWindowHwnd(activeWindow));
                    if (workbookWindow != null)
                    {
                        _logger?.Info(
                            KernelFlickerTracePrefix
                            + " source=TaskPaneRefreshOrchestrationService action=resolve-window-success reason="
                            + (reason ?? string.Empty)
                            + ", workbook="
                            + _formatWorkbookDescriptor(workbook)
                            + ", resolvedWindow="
                            + _formatWindowDescriptor(workbookWindow)
                            + ", resolveAttempt="
                            + (attempt + 1).ToString(CultureInfo.InvariantCulture));
                        return workbookWindow;
                    }

                    if (activeWorkbookMatches && activeWindow != null)
                    {
                        _logger?.Info(
                            KernelFlickerTracePrefix
                            + " source=TaskPaneRefreshOrchestrationService action=resolve-window-success-active-window reason="
                            + (reason ?? string.Empty)
                            + ", workbook="
                            + _formatWorkbookDescriptor(workbook)
                            + ", resolvedWindow="
                            + _formatWindowDescriptor(activeWindow)
                            + ", resolveAttempt="
                            + (attempt + 1).ToString(CultureInfo.InvariantCulture));
                        return activeWindow;
                    }

                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneRefreshOrchestrationService action=resolve-window-retry reason="
                        + (reason ?? string.Empty)
                        + ", workbook="
                        + _formatWorkbookDescriptor(workbook)
                        + ", resolveAttempt="
                        + (attempt + 1).ToString(CultureInfo.InvariantCulture)
                        + ", deferredToRetryCoordinator=true");
                }

                _logger?.Warn(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=resolve-window-failed reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + _formatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _formatActiveState());
                _logger?.Warn(
                    "ResolveWorkbookPaneWindow failed. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + workbookFullName);
                return null;
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string FormatActiveState()
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
    }

}
