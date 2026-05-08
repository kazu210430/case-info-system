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
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly ICasePaneHostBridge _casePaneHostBridge;
        private readonly PendingPaneRefreshRetryService _pendingPaneRefreshRetryService;
        private readonly object _createdCaseDisplaySessionSyncRoot = new object();
        private readonly Dictionary<string, CreatedCaseDisplaySession> _createdCaseDisplaySessions = new Dictionary<string, CreatedCaseDisplaySession>(StringComparer.OrdinalIgnoreCase);

        private readonly List<System.Windows.Forms.Timer> _waitReadyRetryTimers = new List<System.Windows.Forms.Timer>();
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
            _pendingPaneRefreshRetryService = new PendingPaneRefreshRetryService(
                _excelInteropService,
                _workbookSessionService,
                _logger,
                PendingPaneRefreshIntervalMs,
                PendingPaneRefreshMaxAttempts,
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
                + FormatDisplayRequestTraceFields(displayRequest));
            LogWindowActivateDisplayRefreshTriggerStart(displayRequest, reason, workbook, window, refreshAttemptId);
            RefreshPreconditionEvaluationResult preconditionEvaluationResult = RefreshPreconditionEvaluator.Evaluate(reason, workbook, window, _casePaneHostBridge);
            if (!preconditionEvaluationResult.CanRefresh)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action="
                    + preconditionEvaluationResult.SkipActionName
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
                TaskPaneRefreshAttemptResult skippedResult = CompleteVisibilityRecoveryOutcome(
                    reason,
                    workbook,
                    window,
                    TaskPaneRefreshAttemptResult.Skipped(),
                    stopwatch,
                    preconditionEvaluationResult.SkipActionName,
                    null,
                    null);
                skippedResult = CompleteRefreshSourceSelectionOutcome(
                    reason,
                    workbook,
                    window,
                    skippedResult,
                    stopwatch,
                    preconditionEvaluationResult.SkipActionName,
                    null);
                skippedResult = CompleteRebuildFallbackOutcome(
                    reason,
                    workbook,
                    window,
                    skippedResult,
                    stopwatch,
                    preconditionEvaluationResult.SkipActionName,
                    null);
                LogWindowActivateDisplayRefreshTriggerOutcome(
                    displayRequest,
                    reason,
                    workbook,
                    window,
                    skippedResult,
                    stopwatch,
                    refreshAttemptId,
                    preconditionEvaluationResult.SkipActionName);
                return skippedResult;
            }

            RefreshDispatchExecutionResult dispatchExecutionResult = RefreshDispatchShell.Dispatch(
                _taskPaneRefreshCoordinator,
                reason,
                workbook,
                window,
                _getKernelHomeForm,
                _getTaskPaneRefreshSuppressionCount);
            TaskPaneRefreshAttemptResult attemptResult = CompleteVisibilityRecoveryOutcome(
                reason,
                workbook,
                window,
                dispatchExecutionResult.AttemptResult,
                stopwatch,
                "refresh",
                null,
                null);
            attemptResult = CompleteRefreshSourceSelectionOutcome(
                reason,
                workbook,
                window,
                attemptResult,
                stopwatch,
                "refresh",
                null);
            attemptResult = CompleteRebuildFallbackOutcome(
                reason,
                workbook,
                window,
                attemptResult,
                stopwatch,
                "refresh",
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
            LogWindowActivateDisplayRefreshTriggerOutcome(
                displayRequest,
                reason,
                workbook,
                window,
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
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
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
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
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
                ScheduleTaskPaneReadyRetry,
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
            StopWaitReadyRetryTimers();
        }

        private void ScheduleTaskPaneReadyRetry(Excel.Workbook workbook, string reason, int attemptNumber, Action retryAction)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=wait-ready-retry-scheduled reason="
                + (reason ?? string.Empty)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryScheduled=true"
                + ", retryDelayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture)
                + ", delayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
            _logger?.Info(
                "TaskPane wait-ready retry scheduled. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + SafeWorkbookFullName(workbook)
                + ", readyShowReason="
                + (reason ?? string.Empty)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryScheduled=true"
                + ", retryDelayMs="
                + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));

            if (retryAction == null)
            {
                return;
            }

            System.Windows.Forms.Timer retryTimer = new System.Windows.Forms.Timer
            {
                Interval = WorkbookPaneWindowResolveDelayMs
            };

            EventHandler tickHandler = null;
            tickHandler = (sender, args) =>
            {
                retryTimer.Stop();
                retryTimer.Tick -= tickHandler;
                _waitReadyRetryTimers.Remove(retryTimer);
                retryTimer.Dispose();
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=wait-ready-retry-firing reason="
                    + (reason ?? string.Empty)
                    + ", readyShowReason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", attempt="
                    + attemptNumber.ToString(CultureInfo.InvariantCulture)
                    + ", maxAttempts="
                    + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                    + ", retryDelayMs="
                    + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
                _logger?.Info(
                    "TaskPane wait-ready retry firing. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + SafeWorkbookFullName(workbook)
                    + ", readyShowReason="
                    + (reason ?? string.Empty)
                    + ", attempt="
                    + attemptNumber.ToString(CultureInfo.InvariantCulture)
                    + ", maxAttempts="
                    + WorkbookPaneWindowResolveAttempts.ToString(CultureInfo.InvariantCulture)
                    + ", retryDelayMs="
                    + WorkbookPaneWindowResolveDelayMs.ToString(CultureInfo.InvariantCulture));
                retryAction();
            };

            _waitReadyRetryTimers.Add(retryTimer);
            retryTimer.Tick += tickHandler;
            retryTimer.Start();
        }

        private void StopWaitReadyRetryTimers()
        {
            if (_waitReadyRetryTimers.Count == 0)
            {
                return;
            }

            foreach (System.Windows.Forms.Timer retryTimer in _waitReadyRetryTimers.ToArray())
            {
                retryTimer.Stop();
                retryTimer.Dispose();
            }

            _waitReadyRetryTimers.Clear();
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

            VisibilityRecoveryOutcome outcome = BuildVisibilityRecoveryOutcome(
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

        private static VisibilityRecoveryOutcome BuildVisibilityRecoveryOutcome(
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            WorkbookWindowVisibilityEnsureOutcome? ensureStatus = workbookWindowEnsureFacts == null
                ? (WorkbookWindowVisibilityEnsureOutcome?)null
                : workbookWindowEnsureFacts.Outcome;
            VisibilityRecoveryTargetKind targetKind = ResolveVisibilityRecoveryTargetKind(
                workbook,
                inputWindow,
                attemptResult);
            PaneVisibleSource paneVisibleSource = attemptResult.PaneVisibleSource;

            if (!attemptResult.IsRefreshSucceeded)
            {
                if (attemptResult.WasSkipped)
                {
                    return VisibilityRecoveryOutcome.Skipped(
                        "refreshSkipped",
                        isPaneVisible: false,
                        isDisplayCompletable: false,
                        targetKind: targetKind,
                        paneVisibleSource: paneVisibleSource,
                        workbookWindowEnsureStatus: ensureStatus,
                        fullRecoveryAttempted: attemptResult.PreContextRecoveryAttempted,
                        fullRecoverySucceeded: attemptResult.PreContextRecoverySucceeded);
                }

                return VisibilityRecoveryOutcome.Failed(
                    attemptResult.WasContextRejected ? "contextRejected" : "refreshFailed",
                    targetKind,
                    paneVisibleSource,
                    ensureStatus,
                    attemptResult.PreContextRecoveryAttempted,
                    attemptResult.PreContextRecoverySucceeded);
            }

            if (!attemptResult.IsPaneVisible)
            {
                return VisibilityRecoveryOutcome.Failed(
                    "paneVisible=false",
                    targetKind,
                    paneVisibleSource,
                    ensureStatus,
                    attemptResult.PreContextRecoveryAttempted,
                    attemptResult.PreContextRecoverySucceeded);
            }

            string degradedReason = ResolveVisibilityRecoveryDegradedReason(workbookWindowEnsureFacts, attemptResult);
            if (!string.IsNullOrWhiteSpace(degradedReason))
            {
                return VisibilityRecoveryOutcome.Degraded(
                    "paneVisibleWithDegradedRecoveryFacts",
                    targetKind,
                    paneVisibleSource,
                    ensureStatus,
                    attemptResult.PreContextRecoveryAttempted,
                    attemptResult.PreContextRecoverySucceeded,
                    degradedReason);
            }

            if (paneVisibleSource == PaneVisibleSource.AlreadyVisibleHost)
            {
                return VisibilityRecoveryOutcome.Skipped(
                    "alreadyVisible",
                    isPaneVisible: true,
                    isDisplayCompletable: true,
                    targetKind: VisibilityRecoveryTargetKind.AlreadyVisible,
                    paneVisibleSource: paneVisibleSource,
                    workbookWindowEnsureStatus: ensureStatus,
                    fullRecoveryAttempted: attemptResult.PreContextRecoveryAttempted,
                    fullRecoverySucceeded: attemptResult.PreContextRecoverySucceeded);
            }

            string completedReason = ensureStatus == WorkbookWindowVisibilityEnsureOutcome.MadeVisible
                ? "madeVisibleThenShown"
                : "paneVisible";
            if (attemptResult.IsRefreshCompleted)
            {
                completedReason = paneVisibleSource == PaneVisibleSource.ReusedShown
                    ? "reusedShown"
                    : "refreshedShown";
            }

            return VisibilityRecoveryOutcome.Completed(
                completedReason,
                targetKind,
                paneVisibleSource,
                ensureStatus,
                attemptResult.PreContextRecoveryAttempted,
                attemptResult.PreContextRecoverySucceeded);
        }

        private static VisibilityRecoveryTargetKind ResolveVisibilityRecoveryTargetKind(
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            if (attemptResult != null && attemptResult.PaneVisibleSource == PaneVisibleSource.AlreadyVisibleHost)
            {
                return VisibilityRecoveryTargetKind.AlreadyVisible;
            }

            if (workbook == null
                && inputWindow == null
                && attemptResult != null
                && attemptResult.ForegroundWorkbook == null
                && attemptResult.ForegroundWindow == null)
            {
                return attemptResult.ForegroundContext != null
                    || attemptResult.IsRefreshSucceeded
                    || attemptResult.PreContextRecoveryAttempted
                    ? VisibilityRecoveryTargetKind.ActiveWorkbookFallback
                    : VisibilityRecoveryTargetKind.NoKnownTarget;
            }

            if (workbook != null
                || inputWindow != null
                || (attemptResult != null && (attemptResult.ForegroundWorkbook != null || attemptResult.ForegroundWindow != null)))
            {
                return VisibilityRecoveryTargetKind.ExplicitWorkbookWindow;
            }

            return VisibilityRecoveryTargetKind.NoKnownTarget;
        }

        private static string ResolveVisibilityRecoveryDegradedReason(
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            if (workbookWindowEnsureFacts != null)
            {
                switch (workbookWindowEnsureFacts.Outcome)
                {
                    case WorkbookWindowVisibilityEnsureOutcome.WorkbookMissing:
                    case WorkbookWindowVisibilityEnsureOutcome.WindowUnresolved:
                    case WorkbookWindowVisibilityEnsureOutcome.VisibilityReadFailed:
                    case WorkbookWindowVisibilityEnsureOutcome.Failed:
                        return "workbookWindowEnsure=" + workbookWindowEnsureFacts.Outcome.ToString();
                    case WorkbookWindowVisibilityEnsureOutcome.MadeVisible:
                        if (workbookWindowEnsureFacts.VisibleAfterSet != true)
                        {
                            return "workbookWindowEnsureVisibleAfterSet="
                                + (workbookWindowEnsureFacts.VisibleAfterSet.HasValue
                                    ? workbookWindowEnsureFacts.VisibleAfterSet.Value.ToString()
                                    : "null");
                        }

                        break;
                }
            }

            if (attemptResult != null
                && attemptResult.PreContextRecoveryAttempted
                && attemptResult.PreContextRecoverySucceeded == false)
            {
                return "fullRecoveryReturnedFalse";
            }

            return string.Empty;
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
            string detail = FormatVisibilityRecoveryDetails(
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

        private static string FormatVisibilityRecoveryDetails(
            string reason,
            VisibilityRecoveryOutcome outcome,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",visibilityRecoveryStatus=" + outcome.Status.ToString()
                + ",visibilityRecoveryReason=" + outcome.Reason
                + ",visibilityRecoveryTerminal=" + outcome.IsTerminal.ToString()
                + ",visibilityRecoveryDisplayCompletable=" + outcome.IsDisplayCompletable.ToString()
                + ",visibilityRecoveryPaneVisible=" + outcome.IsPaneVisible.ToString()
                + ",visibilityRecoveryTargetKind=" + outcome.TargetKind.ToString()
                + ",visibilityPaneVisibleSource=" + outcome.PaneVisibleSource.ToString()
                + ",visibilityRecoveryDegradedReason=" + outcome.DegradedReason
                + ",refreshSucceeded=" + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ",refreshCompleted=" + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ",preContextFullRecoveryAttempted=" + outcome.FullRecoveryAttempted.ToString()
                + ",preContextFullRecoverySucceeded=" + FormatNullableBool(outcome.FullRecoverySucceeded);
            if (workbookWindowEnsureFacts != null)
            {
                details += ",workbookWindowEnsureStatus=" + workbookWindowEnsureFacts.Outcome.ToString()
                    + ",workbookWindowEnsureHwnd=" + workbookWindowEnsureFacts.WindowHwnd
                    + ",workbookWindowVisibleAfterSet=" + FormatNullableBool(workbookWindowEnsureFacts.VisibleAfterSet);
            }

            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
        }

        private static string FormatNullableBool(bool? value)
        {
            return value.HasValue ? value.Value.ToString() : string.Empty;
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

            RefreshSourceSelectionOutcome outcome = RefreshSourceSelectionOutcome.FromAttemptResult(attemptResult);
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
            string detail = FormatRefreshSourceSelectionDetails(
                reason,
                outcome,
                attemptResult,
                completionSource,
                attemptNumber);
            string statusAction = FormatRefreshSourceSelectionAction(outcome);
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

        private static string FormatRefreshSourceSelectionAction(RefreshSourceSelectionOutcome outcome)
        {
            switch (outcome.Status)
            {
                case RefreshSourceSelectionOutcomeStatus.Selected:
                    return "refresh-source-selected";
                case RefreshSourceSelectionOutcomeStatus.DegradedSelected:
                    return "refresh-source-degraded";
                case RefreshSourceSelectionOutcomeStatus.FallbackSelected:
                    return "refresh-source-fallback";
                case RefreshSourceSelectionOutcomeStatus.RebuildRequired:
                    return "refresh-source-rebuild-required";
                case RefreshSourceSelectionOutcomeStatus.Failed:
                    return "refresh-source-failed";
                case RefreshSourceSelectionOutcomeStatus.NotReached:
                    return "refresh-source-not-reached";
                default:
                    return "refresh-source-unknown";
            }
        }

        private static string FormatRefreshSourceSelectionDetails(
            string reason,
            RefreshSourceSelectionOutcome outcome,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",refreshSourceStatus=" + outcome.Status.ToString()
                + ",selectedSource=" + outcome.SelectedSource.ToString()
                + ",selectionReason=" + outcome.SelectionReason
                + ",fallbackReasons=" + outcome.FallbackReasons
                + ",refreshSourceTerminal=" + outcome.IsTerminal.ToString()
                + ",refreshSourceCanContinue=" + outcome.CanContinueRefresh.ToString()
                + ",cacheFallback=" + outcome.IsCacheFallback.ToString()
                + ",rebuildRequired=" + outcome.IsRebuildRequired.ToString()
                + ",masterListRebuildAttempted=" + outcome.MasterListRebuildAttempted.ToString()
                + ",masterListRebuildSucceeded=" + outcome.MasterListRebuildSucceeded.ToString()
                + ",snapshotTextAvailable=" + outcome.SnapshotTextAvailable.ToString()
                + ",updatedCaseSnapshotCache=" + outcome.UpdatedCaseSnapshotCache.ToString()
                + ",failureReason=" + outcome.FailureReason
                + ",degradedReason=" + outcome.DegradedReason
                + ",refreshSucceeded=" + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ",refreshCompleted=" + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ",paneVisible=" + (attemptResult != null && attemptResult.IsPaneVisible).ToString();
            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
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

            RebuildFallbackOutcome outcome = RebuildFallbackOutcome.FromBuildResult(attemptResult.SnapshotBuildResult);
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
            string detail = FormatRebuildFallbackDetails(
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

        private static string FormatRebuildFallbackDetails(
            string reason,
            RebuildFallbackOutcome outcome,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",rebuildFallbackStatus=" + outcome.Status.ToString()
                + ",rebuildFallbackRequired=" + outcome.IsRequired.ToString()
                + ",rebuildFallbackTerminal=" + outcome.IsTerminal.ToString()
                + ",rebuildFallbackCanContinue=" + outcome.CanContinueRefresh.ToString()
                + ",snapshotSource=" + outcome.SnapshotSource.ToString()
                + ",fallbackReasons=" + outcome.FallbackReasons
                + ",masterListRebuildAttempted=" + outcome.MasterListRebuildAttempted.ToString()
                + ",masterListRebuildSucceeded=" + outcome.MasterListRebuildSucceeded.ToString()
                + ",snapshotTextAvailable=" + outcome.SnapshotTextAvailable.ToString()
                + ",updatedCaseSnapshotCache=" + outcome.UpdatedCaseSnapshotCache.ToString()
                + ",failureReason=" + outcome.FailureReason
                + ",degradedReason=" + outcome.DegradedReason
                + ",outcomeReason=" + outcome.Reason
                + ",refreshSucceeded=" + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ",refreshCompleted=" + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ",paneVisible=" + (attemptResult != null && attemptResult.IsPaneVisible).ToString();
            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
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
                "reason=" + (reason ?? string.Empty)
                + ",foregroundRecoveryStarted=" + foregroundRecoveryStarted.ToString()
                + ",foregroundSkipReason=" + (foregroundSkipReason ?? string.Empty)
                + ",foregroundOutcomeStatus=" + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString()));
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
                "reason=" + (reason ?? string.Empty));
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
                "reason=" + (reason ?? string.Empty)
                + ",recovered=" + (executionResult != null && executionResult.Recovered).ToString()
                + ",foregroundOutcomeStatus="
                + (executionResult != null && executionResult.Recovered
                    ? ForegroundGuaranteeOutcomeStatus.RequiredSucceeded.ToString()
                    : ForegroundGuaranteeOutcomeStatus.RequiredDegraded.ToString()));
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

            Stopwatch stopwatch = Stopwatch.StartNew();
            TaskPaneRefreshAttemptResult attemptResult = CompleteVisibilityRecoveryOutcome(
                reason,
                workbook,
                outcome.WorkbookWindow,
                outcome.RefreshAttemptResult,
                stopwatch,
                "ready-show-attempt",
                outcome.AttemptNumber,
                outcome.WorkbookWindowEnsureFacts);
            attemptResult = CompleteRefreshSourceSelectionOutcome(
                reason,
                workbook,
                outcome.WorkbookWindow,
                attemptResult,
                stopwatch,
                "ready-show-attempt",
                outcome.AttemptNumber);
            attemptResult = CompleteRebuildFallbackOutcome(
                reason,
                workbook,
                outcome.WorkbookWindow,
                attemptResult,
                stopwatch,
                "ready-show-attempt",
                outcome.AttemptNumber);
            attemptResult = CompleteForegroundGuaranteeOutcome(
                reason,
                workbook,
                outcome.WorkbookWindow,
                attemptResult,
                stopwatch);
            TryCompleteCreatedCaseDisplaySession(
                session,
                reason,
                workbook,
                outcome.WorkbookWindow,
                attemptResult,
                "ready-show-attempt",
                outcome.AttemptNumber);
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
            if (!IsCreatedCaseDisplayReason(reason)
                || attemptResult == null
                || !attemptResult.IsRefreshSucceeded
                || !attemptResult.IsPaneVisible
                || attemptResult.VisibilityRecoveryOutcome == null
                || !attemptResult.VisibilityRecoveryOutcome.IsTerminal
                || !attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable
                || !attemptResult.IsForegroundGuaranteeTerminal
                || attemptResult.ForegroundGuaranteeOutcome == null
                || !attemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable)
            {
                return;
            }

            CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);
            if (resolvedSession == null)
            {
                return;
            }

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

            if (!shouldEmit)
            {
                return;
            }

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
                + FormatDisplayRequestDetailFields(displayRequest);
            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

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

        private void LogWindowActivateDisplayRefreshTriggerStart(
            TaskPaneDisplayRequest displayRequest,
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            int refreshAttemptId)
        {
            if (!IsWindowActivateDisplayRequest(displayRequest))
            {
                return;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-start refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", windowActivateDispatchStatus=Dispatched"
                + ", activationAttempt=NotAttempted"
                + ", downstreamRecoveryDelegated=False"
                + ", displayCompletionOutcome=False"
                + ", recoveryOwner=False"
                + ", foregroundGuaranteeOwner=False"
                + ", hiddenExcelOwner=False"
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", inputWindow="
                + FormatWindowDescriptor(window)
                + ", activeState="
                + FormatActiveState()
                + FormatDisplayRequestTraceFields(displayRequest));
        }

        private void LogWindowActivateDisplayRefreshTriggerOutcome(
            TaskPaneDisplayRequest displayRequest,
            string reason,
            Excel.Workbook workbook,
            Excel.Window window,
            TaskPaneRefreshAttemptResult attemptResult,
            Stopwatch stopwatch,
            int refreshAttemptId,
            string completionSource)
        {
            if (!IsWindowActivateDisplayRequest(displayRequest))
            {
                return;
            }

            VisibilityRecoveryOutcome visibilityOutcome = attemptResult == null ? null : attemptResult.VisibilityRecoveryOutcome;
            RefreshSourceSelectionOutcome refreshSourceOutcome = attemptResult == null ? null : attemptResult.RefreshSourceSelectionOutcome;
            RebuildFallbackOutcome rebuildOutcome = attemptResult == null ? null : attemptResult.RebuildFallbackOutcome;
            ForegroundGuaranteeOutcome foregroundOutcome = attemptResult == null ? null : attemptResult.ForegroundGuaranteeOutcome;
            bool downstreamRecoveryDelegated = (attemptResult != null && attemptResult.PreContextRecoveryAttempted)
                || (foregroundOutcome != null && foregroundOutcome.WasExecutionAttempted);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-outcome refreshAttemptId="
                + refreshAttemptId.ToString(CultureInfo.InvariantCulture)
                + ", reason="
                + (reason ?? string.Empty)
                + ", completionSource="
                + (completionSource ?? string.Empty)
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", windowActivateDispatchStatus=Dispatched"
                + ", activationAttempt="
                + (downstreamRecoveryDelegated
                    ? WindowActivateActivationAttempt.Delegated.ToString()
                    : WindowActivateActivationAttempt.NotAttempted.ToString())
                + ", downstreamRecoveryDelegated="
                + downstreamRecoveryDelegated.ToString()
                + ", displayCompletionOutcome=False"
                + ", recoveryOwner=False"
                + ", foregroundGuaranteeOwner=False"
                + ", hiddenExcelOwner=False"
                + ", refreshSucceeded="
                + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ", refreshSkipped="
                + (attemptResult != null && attemptResult.WasSkipped).ToString()
                + ", contextRejected="
                + (attemptResult != null && attemptResult.WasContextRejected).ToString()
                + ", paneVisible="
                + (attemptResult != null && attemptResult.IsPaneVisible).ToString()
                + ", refreshCompleted="
                + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ", visibilityRecoveryStatus="
                + (visibilityOutcome == null ? VisibilityRecoveryOutcomeStatus.Unknown.ToString() : visibilityOutcome.Status.ToString())
                + ", visibilityRecoveryReason="
                + (visibilityOutcome == null ? string.Empty : visibilityOutcome.Reason)
                + ", refreshSourceStatus="
                + (refreshSourceOutcome == null ? RefreshSourceSelectionOutcomeStatus.Unknown.ToString() : refreshSourceOutcome.Status.ToString())
                + ", refreshSourceSelectedSource="
                + (refreshSourceOutcome == null ? TaskPaneSnapshotBuilderService.TaskPaneSnapshotSource.None.ToString() : refreshSourceOutcome.SelectedSource.ToString())
                + ", rebuildFallbackStatus="
                + (rebuildOutcome == null ? RebuildFallbackOutcomeStatus.Unknown.ToString() : rebuildOutcome.Status.ToString())
                + ", foregroundGuaranteeStatus="
                + (foregroundOutcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : foregroundOutcome.Status.ToString())
                + ", foregroundExecutionAttempted="
                + (foregroundOutcome != null && foregroundOutcome.WasExecutionAttempted).ToString()
                + ", preContextFullRecoveryAttempted="
                + (attemptResult != null && attemptResult.PreContextRecoveryAttempted).ToString()
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + FormatDisplayRequestTraceFields(displayRequest));
        }

        private static bool IsWindowActivateDisplayRequest(TaskPaneDisplayRequest displayRequest)
        {
            return displayRequest != null && displayRequest.IsWindowActivateTrigger;
        }

        private static string FormatDisplayRequestTraceFields(TaskPaneDisplayRequest displayRequest)
        {
            if (displayRequest == null)
            {
                return string.Empty;
            }

            return FormatDisplayRequestDetailFields(displayRequest);
        }

        private static string FormatDisplayRequestDetailFields(TaskPaneDisplayRequest displayRequest)
        {
            if (displayRequest == null)
            {
                return string.Empty;
            }

            string details =
                ", displayRequestSource=" + displayRequest.Source.ToString()
                + ", displayRequestRefreshIntent=" + displayRequest.RefreshIntent.ToString()
                + ", displayTriggerReason=" + displayRequest.ToReasonString();
            if (!displayRequest.IsWindowActivateTrigger)
            {
                return details;
            }

            WindowActivateTaskPaneTriggerFacts facts = displayRequest.WindowActivateTriggerFacts;
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

        private static class RefreshPreconditionEvaluator
        {
            internal static RefreshPreconditionEvaluationResult Evaluate(string reason, Excel.Workbook workbook, Excel.Window window, ICasePaneHostBridge casePaneHostBridge)
            {
                if (TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(reason, workbook, window))
                {
                    return RefreshPreconditionEvaluationResult.SkipWorkbookOpenWindowDependentRefresh();
                }

                if (casePaneHostBridge.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window))
                {
                    return RefreshPreconditionEvaluationResult.IgnoreDuringProtection();
                }

                return RefreshPreconditionEvaluationResult.Proceed();
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

        private sealed class RefreshPreconditionEvaluationResult
        {
            private RefreshPreconditionEvaluationResult(bool canRefresh, string skipActionName)
            {
                CanRefresh = canRefresh;
                SkipActionName = skipActionName ?? string.Empty;
            }

            internal bool CanRefresh { get; }

            internal string SkipActionName { get; }

            internal static RefreshPreconditionEvaluationResult Proceed()
            {
                return new RefreshPreconditionEvaluationResult(true, string.Empty);
            }

            internal static RefreshPreconditionEvaluationResult SkipWorkbookOpenWindowDependentRefresh()
            {
                return new RefreshPreconditionEvaluationResult(false, "skip-workbook-open-window-dependent-refresh");
            }

            internal static RefreshPreconditionEvaluationResult IgnoreDuringProtection()
            {
                return new RefreshPreconditionEvaluationResult(false, "ignore-during-protection");
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

    internal sealed class PendingPaneRefreshRetryService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly ExcelInteropService _excelInteropService;
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly Logger _logger;
        private readonly int _pendingPaneRefreshIntervalMs;
        private readonly int _pendingPaneRefreshMaxAttempts;
        private readonly Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> _tryRefreshTaskPane;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Action _stopPendingPaneRefreshTimer;
        private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
        private readonly Func<Excel.Window, string> _formatWindowDescriptor;
        private readonly Func<string> _formatActiveState;
        private readonly Func<Excel.Workbook, string> _safeWorkbookFullName;
        private readonly Func<Excel.Window, string> _safeWindowHwnd;
        private readonly PendingPaneRefreshRetryState _retryState = new PendingPaneRefreshRetryState();

        private System.Windows.Forms.Timer _pendingPaneRefreshTimer;

        internal PendingPaneRefreshRetryService(
            ExcelInteropService excelInteropService,
            WorkbookSessionService workbookSessionService,
            Logger logger,
            int pendingPaneRefreshIntervalMs,
            int pendingPaneRefreshMaxAttempts,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Action stopPendingPaneRefreshTimer,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<Excel.Window, string> formatWindowDescriptor,
            Func<string> formatActiveState,
            Func<Excel.Workbook, string> safeWorkbookFullName,
            Func<Excel.Window, string> safeWindowHwnd)
        {
            _excelInteropService = excelInteropService;
            _workbookSessionService = workbookSessionService;
            _logger = logger;
            _pendingPaneRefreshIntervalMs = pendingPaneRefreshIntervalMs;
            _pendingPaneRefreshMaxAttempts = pendingPaneRefreshMaxAttempts;
            _tryRefreshTaskPane = tryRefreshTaskPane ?? throw new ArgumentNullException(nameof(tryRefreshTaskPane));
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow ?? throw new ArgumentNullException(nameof(resolveWorkbookPaneWindow));
            _stopPendingPaneRefreshTimer = stopPendingPaneRefreshTimer ?? throw new ArgumentNullException(nameof(stopPendingPaneRefreshTimer));
            _formatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
            _formatWindowDescriptor = formatWindowDescriptor ?? throw new ArgumentNullException(nameof(formatWindowDescriptor));
            _formatActiveState = formatActiveState ?? throw new ArgumentNullException(nameof(formatActiveState));
            _safeWorkbookFullName = safeWorkbookFullName ?? throw new ArgumentNullException(nameof(safeWorkbookFullName));
            _safeWindowHwnd = safeWindowHwnd ?? throw new ArgumentNullException(nameof(safeWindowHwnd));
        }

        internal int AttemptsRemaining
        {
            get
            {
                return _retryState.AttemptsRemaining;
            }
        }

        internal void TrackActiveTarget()
        {
            _retryState.TrackActiveTarget();
        }

        internal void TrackWorkbookTarget(string workbookFullName)
        {
            _retryState.TrackWorkbookTarget(workbookFullName);
        }

        internal int BeginRetrySequence(string reason)
        {
            _retryState.BeginRetrySequence(reason, _pendingPaneRefreshMaxAttempts);
            EnsurePendingPaneRefreshTimer();
            _pendingPaneRefreshTimer.Stop();
            _pendingPaneRefreshTimer.Start();
            return _retryState.AttemptsRemaining;
        }

        internal void StopTimer()
        {
            _pendingPaneRefreshTimer?.Stop();
        }

        private void EnsurePendingPaneRefreshTimer()
        {
            if (_pendingPaneRefreshTimer != null)
            {
                return;
            }

            _pendingPaneRefreshTimer = new System.Windows.Forms.Timer();
            _pendingPaneRefreshTimer.Interval = _pendingPaneRefreshIntervalMs;
            _pendingPaneRefreshTimer.Tick += PendingPaneRefreshTimer_Tick;
        }

        private void PendingPaneRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (!_retryState.HasAttemptsRemaining)
            {
                _stopPendingPaneRefreshTimer();
                return;
            }

            Excel.Workbook targetWorkbook = ResolvePendingPaneRefreshWorkbook();
            if (targetWorkbook != null)
            {
                int attemptsRemaining = _retryState.ConsumeAttempt();
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-retry-start reason="
                    + _retryState.Reason
                    + ", workbook="
                    + _formatWorkbookDescriptor(targetWorkbook)
                    + ", attemptsRemaining="
                    + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
                _logger?.Info("TaskPane timer retry start. reason=" + _retryState.Reason + ", workbook=" + _safeWorkbookFullName(targetWorkbook) + ", attemptsRemaining=" + attemptsRemaining.ToString(CultureInfo.InvariantCulture));
                Excel.Window workbookWindow = _resolveWorkbookPaneWindow(targetWorkbook, _retryState.Reason, true);
                bool refreshed = _tryRefreshTaskPane(_retryState.Reason, targetWorkbook, workbookWindow).IsRefreshSucceeded;
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-retry-end reason="
                    + _retryState.Reason
                    + ", workbook="
                    + _formatWorkbookDescriptor(targetWorkbook)
                    + ", resolvedWindow="
                    + _formatWindowDescriptor(workbookWindow)
                    + ", refreshed="
                    + refreshed.ToString());
                _logger?.Info("TaskPane timer retry result. reason=" + _retryState.Reason + ", workbook=" + _safeWorkbookFullName(targetWorkbook) + ", windowHwnd=" + _safeWindowHwnd(workbookWindow) + ", refreshed=" + refreshed.ToString());
                if (refreshed)
                {
                    _stopPendingPaneRefreshTimer();
                }

                return;
            }

            WorkbookContext context = _workbookSessionService == null ? null : _workbookSessionService.ResolveActiveContext("PendingPaneRefresh");
            if (context == null || context.Role != WorkbookRole.Case)
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-stop reason="
                    + _retryState.Reason
                    + ", pendingWorkbookFullName=\""
                    + _retryState.WorkbookFullName
                    + "\", contextRole="
                    + (context == null ? "null" : context.Role.ToString())
                    + ", attemptsRemaining="
                    + _retryState.AttemptsRemaining.ToString(CultureInfo.InvariantCulture));
                _stopPendingPaneRefreshTimer();
                return;
            }

            int fallbackAttemptsRemaining = _retryState.ConsumeAttempt();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-start reason="
                + _retryState.Reason
                + ", pendingWorkbookFullName=\""
                + _retryState.WorkbookFullName
                + "\", contextWorkbook="
                + _formatWorkbookDescriptor(context.Workbook)
                + ", attemptsRemaining="
                + fallbackAttemptsRemaining.ToString(CultureInfo.InvariantCulture)
                + ", activeState="
                + _formatActiveState());
            _logger?.Info(
                "TaskPane timer fallback active CASE context start. reason="
                + _retryState.Reason
                + ", pendingWorkbook="
                + _retryState.WorkbookFullName
                + ", attemptsRemaining="
                + fallbackAttemptsRemaining.ToString(CultureInfo.InvariantCulture));
            bool fallbackRefreshed = _tryRefreshTaskPane(_retryState.Reason, null, null).IsRefreshSucceeded;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneRefreshOrchestrationService action=defer-active-context-fallback-end reason="
                + _retryState.Reason
                + ", pendingWorkbookFullName=\""
                + _retryState.WorkbookFullName
                + "\", contextWorkbook="
                + _formatWorkbookDescriptor(context.Workbook)
                + ", refreshed="
                + fallbackRefreshed.ToString()
                + ", activeState="
                + _formatActiveState());
            _logger?.Info(
                "TaskPane timer fallback active CASE context result. reason="
                + _retryState.Reason
                + ", pendingWorkbook="
                + _retryState.WorkbookFullName
                + ", refreshed="
                + fallbackRefreshed.ToString());
            if (fallbackRefreshed)
            {
                _stopPendingPaneRefreshTimer();
            }
        }

        private Excel.Workbook ResolvePendingPaneRefreshWorkbook()
        {
            if (_excelInteropService == null || string.IsNullOrWhiteSpace(_retryState.WorkbookFullName))
            {
                return null;
            }

            return _excelInteropService.FindOpenWorkbook(_retryState.WorkbookFullName);
        }

        private sealed class PendingPaneRefreshRetryState
        {
            internal int AttemptsRemaining { get; private set; }

            internal string Reason { get; private set; } = string.Empty;

            internal string WorkbookFullName { get; private set; } = string.Empty;

            internal bool HasAttemptsRemaining
            {
                get
                {
                    return AttemptsRemaining > 0;
                }
            }

            internal void TrackActiveTarget()
            {
                WorkbookFullName = string.Empty;
            }

            internal void TrackWorkbookTarget(string workbookFullName)
            {
                WorkbookFullName = workbookFullName ?? string.Empty;
            }

            internal void BeginRetrySequence(string reason, int maxAttempts)
            {
                Reason = reason ?? string.Empty;
                AttemptsRemaining = maxAttempts;
            }

            internal int ConsumeAttempt()
            {
                if (AttemptsRemaining <= 0)
                {
                    return 0;
                }

                AttemptsRemaining--;
                return AttemptsRemaining;
            }
        }
    }
}
