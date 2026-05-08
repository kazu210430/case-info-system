using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum WindowActivateDispatchOutcomeStatus
    {
        Unknown = 0,
        Observed = 1,
        Ignored = 2,
        Deferred = 3,
        Dispatched = 4,
        Failed = 5
    }

    internal enum WindowActivateActivationAttempt
    {
        NotAttempted = 0,
        Delegated = 1,
        Succeeded = 2,
        Failed = 3
    }

    internal sealed class WindowActivateTaskPaneTriggerFacts
    {
        internal WindowActivateTaskPaneTriggerFacts(
            Excel.Workbook workbook,
            Excel.Window window,
            string workbookDescriptor,
            string windowDescriptor,
            string activeState,
            string workbookFullName,
            string windowHwnd,
            string captureOwner)
        {
            Workbook = workbook;
            Window = window;
            WorkbookDescriptor = workbookDescriptor ?? string.Empty;
            WindowDescriptor = windowDescriptor ?? string.Empty;
            ActiveState = activeState ?? string.Empty;
            WorkbookFullName = workbookFullName ?? string.Empty;
            WindowHwnd = windowHwnd ?? string.Empty;
            CaptureOwner = captureOwner ?? string.Empty;
        }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window Window { get; }

        internal string WorkbookDescriptor { get; }

        internal string WindowDescriptor { get; }

        internal string ActiveState { get; }

        internal string WorkbookFullName { get; }

        internal string WindowHwnd { get; }

        internal string CaptureOwner { get; }

        internal bool HasWorkbook
        {
            get { return Workbook != null; }
        }

        internal bool HasWindow
        {
            get { return Window != null; }
        }
    }

    internal sealed class WindowActivateDispatchOutcome
    {
        private WindowActivateDispatchOutcome(
            WindowActivateDispatchOutcomeStatus status,
            WindowActivateTaskPaneTriggerFacts triggerFacts,
            TaskPaneDisplayRequest displayRequest,
            WindowActivateActivationAttempt activationAttempt,
            string outcomeReason,
            bool isTerminal)
        {
            Status = status;
            TriggerFacts = triggerFacts;
            DisplayRequest = displayRequest;
            ActivationAttempt = activationAttempt;
            OutcomeReason = outcomeReason ?? string.Empty;
            IsTerminal = isTerminal;
        }

        internal WindowActivateDispatchOutcomeStatus Status { get; }

        internal WindowActivateTaskPaneTriggerFacts TriggerFacts { get; }

        internal TaskPaneDisplayRequest DisplayRequest { get; }

        internal WindowActivateActivationAttempt ActivationAttempt { get; }

        internal string OutcomeReason { get; }

        internal bool IsTerminal { get; }

        internal bool IsDisplayCompletionOutcome
        {
            get { return false; }
        }

        internal bool IsRecoveryOwner
        {
            get { return false; }
        }

        internal bool IsForegroundGuaranteeOwner
        {
            get { return false; }
        }

        internal bool IsHiddenExcelOwner
        {
            get { return false; }
        }

        internal static WindowActivateDispatchOutcome Observed(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            return new WindowActivateDispatchOutcome(
                WindowActivateDispatchOutcomeStatus.Observed,
                triggerFacts,
                displayRequest: null,
                activationAttempt: WindowActivateActivationAttempt.NotAttempted,
                outcomeReason: "eventObserved",
                isTerminal: false);
        }

        internal static WindowActivateDispatchOutcome Ignored(
            WindowActivateTaskPaneTriggerFacts triggerFacts,
            TaskPaneDisplayRequest displayRequest,
            string outcomeReason)
        {
            return new WindowActivateDispatchOutcome(
                WindowActivateDispatchOutcomeStatus.Ignored,
                triggerFacts,
                displayRequest,
                WindowActivateActivationAttempt.NotAttempted,
                outcomeReason,
                isTerminal: true);
        }

        internal static WindowActivateDispatchOutcome Deferred(
            WindowActivateTaskPaneTriggerFacts triggerFacts,
            TaskPaneDisplayRequest displayRequest,
            string outcomeReason)
        {
            return new WindowActivateDispatchOutcome(
                WindowActivateDispatchOutcomeStatus.Deferred,
                triggerFacts,
                displayRequest,
                WindowActivateActivationAttempt.NotAttempted,
                outcomeReason,
                isTerminal: true);
        }

        internal static WindowActivateDispatchOutcome Dispatched(
            WindowActivateTaskPaneTriggerFacts triggerFacts,
            TaskPaneDisplayRequest displayRequest,
            string outcomeReason)
        {
            return new WindowActivateDispatchOutcome(
                WindowActivateDispatchOutcomeStatus.Dispatched,
                triggerFacts,
                displayRequest,
                WindowActivateActivationAttempt.NotAttempted,
                outcomeReason,
                isTerminal: true);
        }

        internal static WindowActivateDispatchOutcome Failed(
            WindowActivateTaskPaneTriggerFacts triggerFacts,
            string outcomeReason)
        {
            return new WindowActivateDispatchOutcome(
                WindowActivateDispatchOutcomeStatus.Failed,
                triggerFacts,
                displayRequest: null,
                activationAttempt: WindowActivateActivationAttempt.NotAttempted,
                outcomeReason: outcomeReason,
                isTerminal: true);
        }
    }

    internal interface IWindowActivatePanePredicateBridge
    {
        bool ShouldIgnoreDuringCaseProtection(Excel.Workbook workbook, Excel.Window window);
    }

    internal sealed class ThisAddInWindowActivatePanePredicateBridge : IWindowActivatePanePredicateBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInWindowActivatePanePredicateBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public bool ShouldIgnoreDuringCaseProtection(Excel.Workbook workbook, Excel.Window window)
        {
            return _addIn.ShouldIgnoreWindowActivateDuringCaseProtection(workbook, window);
        }
    }

    internal sealed class WindowActivatePaneHandlingService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private readonly IWindowActivatePanePredicateBridge _windowActivatePanePredicateBridge;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Logger _logger;

        internal WindowActivatePaneHandlingService(
            IWindowActivatePanePredicateBridge windowActivatePanePredicateBridge,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane,
            Logger logger = null)
        {
            _windowActivatePanePredicateBridge = windowActivatePanePredicateBridge ?? throw new ArgumentNullException(nameof(windowActivatePanePredicateBridge));
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPane = refreshTaskPane;
            _logger = logger;
        }

        internal WindowActivateDispatchOutcome Handle(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            if (triggerFacts == null)
            {
                WindowActivateDispatchOutcome failedOutcome = WindowActivateDispatchOutcome.Failed(
                    null,
                    "triggerFactsMissing");
                LogDispatchOutcome(failedOutcome);
                return failedOutcome;
            }

            Excel.Workbook workbook = triggerFacts.Workbook;
            Excel.Window window = triggerFacts.Window;
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate(triggerFacts);
            string reason = request.ToReasonString();
            LogDispatchOutcome(WindowActivateDispatchOutcome.Observed(triggerFacts));

            if (_windowActivatePanePredicateBridge.ShouldIgnoreDuringCaseProtection(workbook, window))
            {
                WindowActivateDispatchOutcome ignoredOutcome = WindowActivateDispatchOutcome.Ignored(
                    triggerFacts,
                    request,
                    "caseProtection");
                LogDispatchOutcome(ignoredOutcome);
                return ignoredOutcome;
            }

            _handleExternalWorkbookDetected?.Invoke(workbook, reason);
            if (_shouldSuppressCasePaneRefresh != null && _shouldSuppressCasePaneRefresh(reason, workbook))
            {
                WindowActivateDispatchOutcome deferredOutcome = WindowActivateDispatchOutcome.Deferred(
                    triggerFacts,
                    request,
                    "casePaneRefreshSuppressed");
                LogDispatchOutcome(deferredOutcome);
                return deferredOutcome;
            }

            _refreshTaskPane?.Invoke(request, workbook, window);
            WindowActivateDispatchOutcome dispatchedOutcome = WindowActivateDispatchOutcome.Dispatched(
                triggerFacts,
                request,
                "displayRequestDispatched");
            LogDispatchOutcome(dispatchedOutcome);
            return dispatchedOutcome;
        }

        internal WindowActivateDispatchOutcome Handle(Excel.Workbook workbook, Excel.Window window)
        {
            return Handle(new WindowActivateTaskPaneTriggerFacts(
                workbook,
                window,
                workbookDescriptor: string.Empty,
                windowDescriptor: string.Empty,
                activeState: string.Empty,
                workbookFullName: string.Empty,
                windowHwnd: SafeWindowHwnd(window),
                captureOwner: "WindowActivatePaneHandlingService.Handle.Legacy"));
        }

        private void LogDispatchOutcome(WindowActivateDispatchOutcome outcome)
        {
            if (outcome == null)
            {
                return;
            }

            WindowActivateTaskPaneTriggerFacts facts = outcome.TriggerFacts;
            string action = ResolveDispatchAction(outcome.Status);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WindowActivatePaneHandlingService action="
                + action
                + " reason=WindowActivate"
                + ", triggerRole=TaskPaneDisplayRefreshTrigger"
                + ", dispatchStatus="
                + outcome.Status.ToString()
                + ", dispatchTerminal="
                + outcome.IsTerminal.ToString()
                + ", dispatchReason="
                + outcome.OutcomeReason
                + ", requestSource="
                + (outcome.DisplayRequest == null ? string.Empty : outcome.DisplayRequest.Source.ToString())
                + ", activationAttempt="
                + outcome.ActivationAttempt.ToString()
                + ", displayCompletionOutcome="
                + outcome.IsDisplayCompletionOutcome.ToString()
                + ", recoveryOwner="
                + outcome.IsRecoveryOwner.ToString()
                + ", foregroundGuaranteeOwner="
                + outcome.IsForegroundGuaranteeOwner.ToString()
                + ", hiddenExcelOwner="
                + outcome.IsHiddenExcelOwner.ToString()
                + ", workbookNull="
                + (facts == null || facts.Workbook == null).ToString()
                + ", windowHwnd="
                + (facts == null ? string.Empty : facts.WindowHwnd)
                + ", captureOwner="
                + (facts == null ? string.Empty : facts.CaptureOwner)
                + ", activeState="
                + (facts == null ? string.Empty : facts.ActiveState));
        }

        private static string ResolveDispatchAction(WindowActivateDispatchOutcomeStatus status)
        {
            switch (status)
            {
                case WindowActivateDispatchOutcomeStatus.Observed:
                    return "display-refresh-trigger-observed";
                case WindowActivateDispatchOutcomeStatus.Ignored:
                    return "display-refresh-trigger-ignored";
                case WindowActivateDispatchOutcomeStatus.Deferred:
                    return "display-refresh-trigger-deferred";
                case WindowActivateDispatchOutcomeStatus.Dispatched:
                    return "display-refresh-trigger-dispatched";
                case WindowActivateDispatchOutcomeStatus.Failed:
                    return "display-refresh-trigger-failed";
                default:
                    return "display-refresh-trigger-unknown";
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            return window == null ? "0" : window.Hwnd.ToString();
        }
    }
}
