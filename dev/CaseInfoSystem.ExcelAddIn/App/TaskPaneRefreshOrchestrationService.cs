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
            TaskPaneRefreshPreconditionDecision preconditionDecision = TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(
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
                TaskPaneRefreshAttemptResult skippedResult = CompleteNormalizedOutcomeChain(
                    reason,
                    workbook,
                    window,
                    TaskPaneRefreshAttemptResult.Skipped(preconditionDecision.SkipActionName),
                    stopwatch,
                    preconditionDecision.SkipActionName,
                    null,
                    null);
                _windowActivateDownstreamObservation.LogOutcome(
                    displayRequest,
                    reason,
                    skippedResult,
                    stopwatch,
                    refreshAttemptId,
                    preconditionDecision.SkipActionName);
                return skippedResult;
            }

            RefreshDispatchExecutionResult DispatchTaskPaneRefreshRoute()
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

            RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute();
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

        internal void ScheduleActiveTaskPaneRefresh(string reason)
        {
            _pendingPaneRefreshRetryService.TrackActiveTarget();
            if (IsTaskPaneRefreshSucceeded(reason, null, null))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-immediate-success reason="
                    + (reason ?? string.Empty)
                    + ", target=active");
                StopPendingPaneRefreshTimer();
                return;
            }

            int attemptsRemaining = _pendingPaneRefreshRetryService.BeginRetrySequence(reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (reason ?? string.Empty)
                + ", target=active"
                + ", attempts="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
        }

        internal void ScheduleWorkbookTaskPaneRefresh(Excel.Workbook workbook, string reason)
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
            if (TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window: null))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=skip-workbook-open-defer reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveState());
                return;
            }

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

            if (TryRefreshTaskPane(reason, workbook, workbookWindow).IsRefreshSucceeded)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-immediate-success reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook));
                _logger?.Info("TaskPane timer fallback immediate refresh succeeded. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook));
                StopPendingPaneRefreshTimer();
                return;
            }

            int attemptsRemaining = _pendingPaneRefreshRetryService.BeginRetrySequence(reason);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-scheduled reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", attempts="
                + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
            _logger?.Info("TaskPane timer fallback scheduled. reason=" + (reason ?? string.Empty) + ", workbook=" + SafeWorkbookFullName(workbook) + ", attempts=" + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
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
            attemptResult = CompleteVisibilityRecoveryOutcome(
                reason,
                workbook,
                inputWindow,
                attemptResult,
                stopwatch,
                completionSource,
                attemptNumber,
                workbookWindowEnsureFacts);
            attemptResult = CompleteRefreshSourceSelectionOutcome(
                reason,
                workbook,
                inputWindow,
                attemptResult,
                stopwatch,
                completionSource,
                attemptNumber);
            return CompleteRebuildFallbackOutcome(
                reason,
                workbook,
                inputWindow,
                attemptResult,
                stopwatch,
                completionSource,
                attemptNumber);
        }

        private TaskPaneRefreshAttemptResult CompleteVisibilityRecoveryOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            if (attemptResult == null)
            {
                return null;
            }

            VisibilityRecoveryOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(
                workbook,
                inputWindow,
                attemptResult,
                workbookWindowEnsureFacts);
            LogVisibilityRecoveryOutcome(
                reason,
                workbook,
                inputWindow,
                attemptResult,
                outcome,
                stopwatch,
                completionSource,
                attemptNumber,
                workbookWindowEnsureFacts);
            return attemptResult.WithVisibilityRecoveryOutcome(outcome);
        }

        private void LogVisibilityRecoveryOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            VisibilityRecoveryOutcome outcome,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            if (!IsCreatedCaseDisplayReason(reason) || outcome == null)
            {
                return;
            }

            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            Excel.Window observedWindow = context == null ? inputWindow : context.Window;
            string detail = TaskPaneNormalizedOutcomeMapper.FormatVisibilityRecoveryDetails(
                reason,
                outcome,
                attemptResult,
                completionSource,
                attemptNumber,
                workbookWindowEnsureFacts);
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
                observedWindow,
                "visibility-recovery-decision",
                "TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                detail);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                observedWindow,
                "visibility-recovery-" + outcome.Status.ToString().ToLowerInvariant(),
                "TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                detail);
        }

        private TaskPaneRefreshAttemptResult CompleteRefreshSourceSelectionOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber)
        {
            if (attemptResult == null)
            {
                return null;
            }

            RefreshSourceSelectionOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(attemptResult);
            LogRefreshSourceSelectionOutcome(
                reason,
                workbook,
                inputWindow,
                attemptResult,
                outcome,
                stopwatch,
                completionSource,
                attemptNumber);
            return attemptResult.WithRefreshSourceSelectionOutcome(outcome);
        }

        private void LogRefreshSourceSelectionOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            RefreshSourceSelectionOutcome outcome,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber)
        {
            if (!IsCreatedCaseDisplayReason(reason) || outcome == null)
            {
                return;
            }

            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            Excel.Window observedWindow = context == null ? inputWindow : context.Window;
            string detail = TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionDetails(
                reason,
                outcome,
                attemptResult,
                completionSource,
                attemptNumber);
            string statusAction = TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(outcome);
            LogRefreshSourceSelectionTrace(
                reason,
                workbook,
                context,
                observedWindow,
                outcome,
                stopwatch,
                statusAction,
                detail);

            if (outcome.IsRebuildRequired && !string.Equals(statusAction, "refresh-source-rebuild-required", StringComparison.OrdinalIgnoreCase))
            {
                LogRefreshSourceSelectionTrace(
                    reason,
                    workbook,
                    context,
                    observedWindow,
                    outcome,
                    stopwatch,
                    "refresh-source-rebuild-required",
                    detail);
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

        private TaskPaneRefreshAttemptResult CompleteRebuildFallbackOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber)
        {
            if (attemptResult == null)
            {
                return null;
            }

            RebuildFallbackOutcome outcome = TaskPaneNormalizedOutcomeMapper.BuildRebuildFallbackOutcome(attemptResult);
            LogRebuildFallbackOutcome(
                reason,
                workbook,
                inputWindow,
                attemptResult,
                outcome,
                stopwatch,
                completionSource,
                attemptNumber);
            return attemptResult.WithRebuildFallbackOutcome(outcome);
        }

        private void LogRebuildFallbackOutcome(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            RebuildFallbackOutcome outcome,
            Stopwatch stopwatch,
            string completionSource,
            int? attemptNumber)
        {
            if (!IsCreatedCaseDisplayReason(reason) || outcome == null)
            {
                return;
            }

            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            Excel.Window observedWindow = context == null ? inputWindow : context.Window;
            string detail = TaskPaneNormalizedOutcomeMapper.FormatRebuildFallbackDetails(
                reason,
                outcome,
                attemptResult,
                completionSource,
                attemptNumber);
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
                    observedWindow,
                    "rebuild-fallback-required",
                    "TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome",
                    ResolveObservedWorkbookPath(context, workbook),
                    detail);
            }

            string statusAction = "rebuild-fallback-" + outcome.Status.ToString().ToLowerInvariant();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action="
                + statusAction
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
                observedWindow,
                statusAction,
                "TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                detail);
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

            ForegroundGuaranteeOutcome existingOutcome = attemptResult.ForegroundGuaranteeOutcome;
            if (existingOutcome != null
                && existingOutcome.Status == ForegroundGuaranteeOutcomeStatus.SkippedAlreadyVisible)
            {
                LogForegroundGuaranteeDecision(
                    reason,
                    workbook,
                    inputWindow,
                    attemptResult,
                    existingOutcome,
                    foregroundRecoveryStarted: false,
                    foregroundSkipReason: existingOutcome.Reason,
                    stopwatch: stopwatch);
                return attemptResult;
            }

            if (!attemptResult.IsRefreshSucceeded || !attemptResult.IsPaneVisible)
            {
                ForegroundGuaranteeOutcome skippedOutcome = attemptResult.WasSkipped || attemptResult.WasContextRejected
                    ? attemptResult.ForegroundGuaranteeOutcome
                    : ForegroundGuaranteeOutcome.NotRequired("refreshSucceeded=false");
                LogForegroundGuaranteeDecision(
                    reason,
                    workbook,
                    inputWindow,
                    attemptResult,
                    skippedOutcome,
                    foregroundRecoveryStarted: false,
                    foregroundSkipReason: skippedOutcome == null ? string.Empty : skippedOutcome.Reason,
                    stopwatch: stopwatch);
                return attemptResult.WithForegroundGuaranteeOutcome(skippedOutcome);
            }

            bool foregroundRecoveryStarted = attemptResult.IsRefreshCompleted
                && attemptResult.ForegroundWindow != null
                && attemptResult.IsForegroundRecoveryServiceAvailable;
            string foregroundSkipReason = string.Empty;
            if (!attemptResult.IsRefreshCompleted)
            {
                foregroundSkipReason = "refreshCompleted=false";
            }
            else if (attemptResult.ForegroundWindow == null)
            {
                foregroundSkipReason = "window=null";
            }
            else if (!attemptResult.IsForegroundRecoveryServiceAvailable)
            {
                foregroundSkipReason = "recoveryService=null";
            }

            if (!foregroundRecoveryStarted)
            {
                ForegroundGuaranteeOutcome notRequiredOutcome = ForegroundGuaranteeOutcome.NotRequired(foregroundSkipReason);
                LogForegroundGuaranteeDecision(
                    reason,
                    workbook,
                    inputWindow,
                    attemptResult,
                    notRequiredOutcome,
                    foregroundRecoveryStarted: false,
                    foregroundSkipReason: foregroundSkipReason,
                    stopwatch: stopwatch);
                return attemptResult.WithForegroundGuaranteeOutcome(notRequiredOutcome);
            }

            ForegroundGuaranteeTargetKind targetKind = attemptResult.ForegroundWorkbook == null
                ? ForegroundGuaranteeTargetKind.ActiveWorkbookFallback
                : ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow;
            ForegroundGuaranteeOutcome requiredOutcome = ExecuteForegroundGuaranteeAndBuildOutcome(
                reason,
                workbook,
                attemptResult,
                targetKind,
                stopwatch);
            return attemptResult.WithForegroundGuaranteeOutcome(requiredOutcome);
        }

        private ForegroundGuaranteeOutcome ExecuteForegroundGuaranteeAndBuildOutcome(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeTargetKind targetKind,
            Stopwatch stopwatch)
        {
            ForegroundGuaranteeOutcome pendingOutcome = ForegroundGuaranteeOutcome.Unknown("executionPending");
            LogForegroundGuaranteeDecision(
                reason,
                workbook,
                attemptResult.ForegroundWindow,
                attemptResult,
                pendingOutcome,
                foregroundRecoveryStarted: true,
                foregroundSkipReason: string.Empty,
                stopwatch: stopwatch);
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

            return ClassifyRequiredForegroundExecutionOutcome(targetKind, executionResult);
        }

        private static ForegroundGuaranteeOutcome ClassifyRequiredForegroundExecutionOutcome(
            ForegroundGuaranteeTargetKind targetKind,
            ForegroundGuaranteeExecutionResult executionResult)
        {
            return executionResult.ExecutionAttempted && executionResult.Recovered
                ? ForegroundGuaranteeOutcome.RequiredSucceeded(targetKind, "foregroundRecoverySucceeded")
                : ForegroundGuaranteeOutcome.RequiredDegraded(targetKind, "foregroundRecoveryReturnedFalse");
        }

        private void LogForegroundGuaranteeDecision(
            string reason,
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeOutcome outcome,
            bool foregroundRecoveryStarted,
            string foregroundSkipReason,
            Stopwatch stopwatch)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            Excel.Window resolvedWindow = attemptResult == null || attemptResult.ForegroundWindow == null
                ? inputWindow
                : attemptResult.ForegroundWindow;
            bool recoveryServicePresent = attemptResult != null && attemptResult.IsForegroundRecoveryServiceAvailable;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=foreground-recovery-decision reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", refreshSucceeded="
                + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ", resolvedWindowPresent="
                + (resolvedWindow != null).ToString()
                + ", recoveryServicePresent="
                + recoveryServicePresent.ToString()
                + ", foregroundRecoveryStarted="
                + foregroundRecoveryStarted.ToString()
                + ", foregroundRecoverySkipped="
                + (!foregroundRecoveryStarted).ToString()
                + ", foregroundSkipReason="
                + (foregroundSkipReason ?? string.Empty)
                + ", foregroundOutcomeStatus="
                + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString())
                + ", foregroundOutcomeDisplayCompletable="
                + (outcome != null && outcome.IsDisplayCompletable).ToString()
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString()
                + FormatObservationCorrelationFields(context, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                context == null ? inputWindow : context.Window,
                "foreground-recovery-decision",
                "TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                BuildForegroundRecoveryDecisionDetails(
                    reason,
                    foregroundRecoveryStarted,
                    foregroundSkipReason,
                    outcome));
        }

        private void LogFinalForegroundGuaranteeStarted(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=final-foreground-guarantee-start reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString()
                + FormatObservationCorrelationFields(context, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                context == null ? null : context.Window,
                "final-foreground-guarantee-started",
                "TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                BuildFinalForegroundGuaranteeStartedDetails(reason));
        }

        private void LogFinalForegroundGuaranteeCompleted(
            string reason,
            Excel.Workbook workbook,
            TaskPaneRefreshAttemptResult attemptResult,
            ForegroundGuaranteeExecutionResult executionResult,
            Stopwatch stopwatch)
        {
            WorkbookContext context = attemptResult == null ? null : attemptResult.ForegroundContext;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=final-foreground-guarantee-end reason="
                + (reason ?? string.Empty)
                + ", context="
                + FormatContextDescriptor(context)
                + ", recovered="
                + (executionResult != null && executionResult.Recovered).ToString()
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString()
                + FormatObservationCorrelationFields(context, workbook));
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                context == null ? workbook : context.Workbook,
                context == null ? null : context.Window,
                "final-foreground-guarantee-completed",
                "TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome",
                ResolveObservedWorkbookPath(context, workbook),
                BuildFinalForegroundGuaranteeCompletedDetails(reason, executionResult));
        }

        private static string BuildForegroundRecoveryDecisionDetails(
            string reason,
            bool foregroundRecoveryStarted,
            string foregroundSkipReason,
            ForegroundGuaranteeOutcome outcome)
        {
            return "reason=" + (reason ?? string.Empty)
                + ",foregroundRecoveryStarted=" + foregroundRecoveryStarted.ToString()
                + ",foregroundSkipReason=" + (foregroundSkipReason ?? string.Empty)
                + ",foregroundOutcomeStatus=" + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString());
        }

        private static string BuildFinalForegroundGuaranteeStartedDetails(string reason)
        {
            return "reason=" + (reason ?? string.Empty);
        }

        private static string BuildFinalForegroundGuaranteeCompletedDetails(
            string reason,
            ForegroundGuaranteeExecutionResult executionResult)
        {
            return "reason=" + (reason ?? string.Empty)
                + ",recovered=" + (executionResult != null && executionResult.Recovered).ToString()
                + ",foregroundOutcomeStatus="
                + (executionResult != null && executionResult.Recovered
                    ? ForegroundGuaranteeOutcomeStatus.RequiredSucceeded.ToString()
                    : ForegroundGuaranteeOutcomeStatus.RequiredDegraded.ToString());
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
                EvaluateCreatedCaseDisplayCompletionDecision(reason, attemptResult);
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

            string details = BuildCaseDisplayCompletedDetailsPayload(
                reason,
                resolvedSession,
                attemptResult,
                completionSource,
                attemptNumber,
                displayRequest);

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=case-display-completed sessionId="
                + resolvedSession.SessionId
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", window="
                + FormatWindowDescriptor(window)
                + ", completion="
                + attemptResult.CompletionBasis);
            NewCaseVisibilityObservation.Log(
                _logger,
                _excelInteropService,
                null,
                workbook,
                window,
                "case-display-completed",
                "TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession",
                resolvedSession.WorkbookFullName,
                details);
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

        private static string BuildCaseDisplayCompletedDetailsPayload(
            string reason,
            CreatedCaseDisplaySession resolvedSession,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            TaskPaneDisplayRequest displayRequest)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",sessionId=" + resolvedSession.SessionId
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",completion=" + attemptResult.CompletionBasis
                + ",paneVisible=" + attemptResult.IsPaneVisible.ToString()
                + ",visibilityRecoveryStatus=" + attemptResult.VisibilityRecoveryOutcome.Status.ToString()
                + ",visibilityRecoveryDisplayCompletable=" + attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable.ToString()
                + ",visibilityRecoveryPaneVisible=" + attemptResult.VisibilityRecoveryOutcome.IsPaneVisible.ToString()
                + ",visibilityRecoveryTargetKind=" + attemptResult.VisibilityRecoveryOutcome.TargetKind.ToString()
                + ",visibilityPaneVisibleSource=" + attemptResult.VisibilityRecoveryOutcome.PaneVisibleSource.ToString()
                + ",visibilityRecoveryReason=" + attemptResult.VisibilityRecoveryOutcome.Reason
                + ",visibilityRecoveryDegradedReason=" + attemptResult.VisibilityRecoveryOutcome.DegradedReason
                + ",refreshSourceStatus=" + attemptResult.RefreshSourceSelectionOutcome.Status.ToString()
                + ",refreshSourceSelectedSource=" + attemptResult.RefreshSourceSelectionOutcome.SelectedSource.ToString()
                + ",refreshSourceSelectionReason=" + attemptResult.RefreshSourceSelectionOutcome.SelectionReason
                + ",refreshSourceFallbackReasons=" + attemptResult.RefreshSourceSelectionOutcome.FallbackReasons
                + ",refreshSourceCacheFallback=" + attemptResult.RefreshSourceSelectionOutcome.IsCacheFallback.ToString()
                + ",refreshSourceRebuildRequired=" + attemptResult.RefreshSourceSelectionOutcome.IsRebuildRequired.ToString()
                + ",refreshSourceCanContinue=" + attemptResult.RefreshSourceSelectionOutcome.CanContinueRefresh.ToString()
                + ",refreshSourceFailureReason=" + attemptResult.RefreshSourceSelectionOutcome.FailureReason
                + ",refreshSourceDegradedReason=" + attemptResult.RefreshSourceSelectionOutcome.DegradedReason
                + ",rebuildFallbackStatus=" + attemptResult.RebuildFallbackOutcome.Status.ToString()
                + ",rebuildFallbackRequired=" + attemptResult.RebuildFallbackOutcome.IsRequired.ToString()
                + ",rebuildFallbackCanContinue=" + attemptResult.RebuildFallbackOutcome.CanContinueRefresh.ToString()
                + ",rebuildFallbackSnapshotSource=" + attemptResult.RebuildFallbackOutcome.SnapshotSource.ToString()
                + ",rebuildFallbackReasons=" + attemptResult.RebuildFallbackOutcome.FallbackReasons
                + ",rebuildFallbackFailureReason=" + attemptResult.RebuildFallbackOutcome.FailureReason
                + ",rebuildFallbackDegradedReason=" + attemptResult.RebuildFallbackOutcome.DegradedReason
                + ",refreshCompleted=" + attemptResult.IsRefreshCompleted.ToString()
                + ",foregroundGuaranteeTerminal=" + attemptResult.IsForegroundGuaranteeTerminal.ToString()
                + ",foregroundGuaranteeRequired=" + attemptResult.WasForegroundGuaranteeRequired.ToString()
                + ",foregroundGuaranteeStatus=" + attemptResult.ForegroundGuaranteeOutcome.Status.ToString()
                + ",foregroundGuaranteeDisplayCompletable=" + attemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable.ToString()
                + ",foregroundGuaranteeExecutionAttempted=" + attemptResult.ForegroundGuaranteeOutcome.WasExecutionAttempted.ToString()
                + ",foregroundGuaranteeTargetKind=" + attemptResult.ForegroundGuaranteeOutcome.TargetKind.ToString()
                + ",foregroundRecoverySucceeded="
                + (attemptResult.ForegroundGuaranteeOutcome.RecoverySucceeded.HasValue
                    ? attemptResult.ForegroundGuaranteeOutcome.RecoverySucceeded.Value.ToString()
                    : string.Empty)
                + ",foregroundOutcomeReason=" + attemptResult.ForegroundGuaranteeOutcome.Reason
                + WindowActivateDownstreamObservation.FormatDisplayRequestTraceFields(displayRequest);
            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
        }

        private static CreatedCaseDisplayCompletionDecision EvaluateCreatedCaseDisplayCompletionDecision(
            string reason,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            if (!IsCreatedCaseDisplayReason(reason)
                || attemptResult == null)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            if (!attemptResult.IsRefreshSucceeded)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            if (!attemptResult.IsPaneVisible)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            if (attemptResult.VisibilityRecoveryOutcome == null)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            if (!attemptResult.VisibilityRecoveryOutcome.IsTerminal)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            if (!attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            if (!IsForegroundDisplayCompletableTerminalInput(attemptResult.ForegroundGuaranteeOutcome))
            {
                return CreatedCaseDisplayCompletionDecision.Blocked();
            }

            return CreatedCaseDisplayCompletionDecision.Allowed();
        }

        private static bool IsForegroundDisplayCompletableTerminalInput(ForegroundGuaranteeOutcome outcome)
        {
            return outcome != null
                && outcome.IsTerminal
                && outcome.IsDisplayCompletable;
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
            return string.Equals(reason, NewCaseDefaultTimingLogHelper.PostReleaseReason, StringComparison.OrdinalIgnoreCase);
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

        private struct CreatedCaseDisplayCompletionDecision
        {
            private CreatedCaseDisplayCompletionDecision(bool canComplete)
            {
                CanComplete = canComplete;
            }

            internal bool CanComplete { get; }

            internal static CreatedCaseDisplayCompletionDecision Allowed()
            {
                return new CreatedCaseDisplayCompletionDecision(true);
            }

            internal static CreatedCaseDisplayCompletionDecision Blocked()
            {
                return new CreatedCaseDisplayCompletionDecision(false);
            }
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
