using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum TaskPaneRefreshCompletionDecisionStatus
    {
        Blocked = 0,
        ReadyForSession = 1,
    }

    internal enum TaskPaneRefreshCompletionResultStatus
    {
        Blocked = 0,
        SessionMissing = 1,
        ReadyToEmit = 2,
    }

    internal sealed class TaskPaneRefreshCompletionDecisionService
    {
        internal CreatedCaseDisplayCompletionDecision DecideCreatedCaseDisplayCompletion(
            TaskPaneRefreshCompletionContext context)
        {
            TaskPaneRefreshCompletionMaterial material = TaskPaneRefreshCompletionMaterial.FromContext(context);
            if (!material.IsCreatedCaseDisplayReason)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("reasonNotCreatedCaseDisplay", material);
            }

            if (!material.HasAttemptResult)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("attemptResult=null", material);
            }

            if (!material.IsRefreshSucceeded)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("refreshSucceeded=false", material);
            }

            if (!material.IsPaneVisible)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("paneVisible=false", material);
            }

            if (!material.HasVisibilityRecoveryOutcome)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("visibilityRecoveryOutcome=null", material);
            }

            if (!material.IsVisibilityRecoveryTerminal)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("visibilityRecoveryTerminal=false", material);
            }

            if (!material.IsVisibilityRecoveryDisplayCompletable)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("visibilityRecoveryDisplayCompletable=false", material);
            }

            if (!material.IsForegroundDisplayCompletableTerminalInput)
            {
                return CreatedCaseDisplayCompletionDecision.Blocked("foregroundGuaranteeDisplayCompletable=false", material);
            }

            return CreatedCaseDisplayCompletionDecision.Allowed(material);
        }

        internal TaskPaneRefreshCompletionResult ClassifyCreatedCaseDisplayCompletionResult(
            TaskPaneRefreshCompletionContext context,
            CreatedCaseDisplayCompletionDecision decision,
            CreatedCaseDisplaySessionSnapshot sessionSnapshot)
        {
            if (!decision.CanComplete)
            {
                return TaskPaneRefreshCompletionResult.Blocked(context, decision.BlockedReason);
            }

            if (sessionSnapshot == null)
            {
                return TaskPaneRefreshCompletionResult.SessionMissing(context, "session=null");
            }

            return TaskPaneRefreshCompletionResult.ReadyToEmit(context, sessionSnapshot);
        }

        internal static bool IsForegroundDisplayCompletableTerminalInput(ForegroundGuaranteeOutcome outcome)
        {
            return outcome != null
                && outcome.IsTerminal
                && outcome.IsDisplayCompletable;
        }
    }

    internal sealed class TaskPaneRefreshCompletionContextInput
    {
        internal TaskPaneRefreshCompletionContextInput(
            string reason,
            bool isCreatedCaseDisplayReason,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            TaskPaneDisplayRequest displayRequest,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            Reason = reason ?? string.Empty;
            IsCreatedCaseDisplayReason = isCreatedCaseDisplayReason;
            AttemptResult = attemptResult;
            CompletionSource = completionSource ?? string.Empty;
            AttemptNumber = attemptNumber;
            DisplayRequest = displayRequest;
            Workbook = workbook;
            Window = window;
        }

        internal string Reason { get; }

        internal bool IsCreatedCaseDisplayReason { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string CompletionSource { get; }

        internal int? AttemptNumber { get; }

        internal TaskPaneDisplayRequest DisplayRequest { get; }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window Window { get; }
    }

    internal sealed class TaskPaneRefreshCompletionContext
    {
        private TaskPaneRefreshCompletionContext(TaskPaneRefreshCompletionContextInput input)
        {
            Reason = input.Reason;
            IsCreatedCaseDisplayReason = input.IsCreatedCaseDisplayReason;
            AttemptResult = input.AttemptResult;
            CompletionSource = input.CompletionSource;
            AttemptNumber = input.AttemptNumber;
            DisplayRequest = input.DisplayRequest;
            Workbook = input.Workbook;
            Window = input.Window;
        }

        internal string Reason { get; }

        internal bool IsCreatedCaseDisplayReason { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string CompletionSource { get; }

        internal int? AttemptNumber { get; }

        internal TaskPaneDisplayRequest DisplayRequest { get; }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window Window { get; }

        internal static TaskPaneRefreshCompletionContext FromInput(TaskPaneRefreshCompletionContextInput input)
        {
            if (input == null)
            {
                throw new ArgumentNullException(nameof(input));
            }

            return new TaskPaneRefreshCompletionContext(input);
        }
    }

    internal sealed class TaskPaneRefreshCompletionMaterial
    {
        private TaskPaneRefreshCompletionMaterial(
            bool isCreatedCaseDisplayReason,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            IsCreatedCaseDisplayReason = isCreatedCaseDisplayReason;
            AttemptResult = attemptResult;
            HasAttemptResult = attemptResult != null;
            IsRefreshSucceeded = attemptResult != null && attemptResult.IsRefreshSucceeded;
            IsPaneVisible = attemptResult != null && attemptResult.IsPaneVisible;
            VisibilityRecoveryOutcome = attemptResult == null ? null : attemptResult.VisibilityRecoveryOutcome;
            ForegroundGuaranteeOutcome = attemptResult == null ? null : attemptResult.ForegroundGuaranteeOutcome;
            HasVisibilityRecoveryOutcome = VisibilityRecoveryOutcome != null;
            IsVisibilityRecoveryTerminal = VisibilityRecoveryOutcome != null && VisibilityRecoveryOutcome.IsTerminal;
            IsVisibilityRecoveryDisplayCompletable = VisibilityRecoveryOutcome != null && VisibilityRecoveryOutcome.IsDisplayCompletable;
            IsForegroundDisplayCompletableTerminalInput = TaskPaneRefreshCompletionDecisionService.IsForegroundDisplayCompletableTerminalInput(ForegroundGuaranteeOutcome);
        }

        internal bool IsCreatedCaseDisplayReason { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal bool HasAttemptResult { get; }

        internal bool IsRefreshSucceeded { get; }

        internal bool IsPaneVisible { get; }

        internal VisibilityRecoveryOutcome VisibilityRecoveryOutcome { get; }

        internal ForegroundGuaranteeOutcome ForegroundGuaranteeOutcome { get; }

        internal bool HasVisibilityRecoveryOutcome { get; }

        internal bool IsVisibilityRecoveryTerminal { get; }

        internal bool IsVisibilityRecoveryDisplayCompletable { get; }

        internal bool IsForegroundDisplayCompletableTerminalInput { get; }

        internal static TaskPaneRefreshCompletionMaterial FromContext(TaskPaneRefreshCompletionContext context)
        {
            return new TaskPaneRefreshCompletionMaterial(
                context != null && context.IsCreatedCaseDisplayReason,
                context == null ? null : context.AttemptResult);
        }
    }

    internal struct CreatedCaseDisplayCompletionDecision
    {
        private CreatedCaseDisplayCompletionDecision(
            TaskPaneRefreshCompletionDecisionStatus status,
            string blockedReason,
            TaskPaneRefreshCompletionMaterial material)
        {
            Status = status;
            BlockedReason = blockedReason ?? string.Empty;
            Material = material;
        }

        internal TaskPaneRefreshCompletionDecisionStatus Status { get; }

        internal bool CanComplete
        {
            get
            {
                return Status == TaskPaneRefreshCompletionDecisionStatus.ReadyForSession;
            }
        }

        internal bool ShouldResolveSession
        {
            get
            {
                return CanComplete;
            }
        }

        internal string BlockedReason { get; }

        internal TaskPaneRefreshCompletionMaterial Material { get; }

        internal static CreatedCaseDisplayCompletionDecision Allowed(TaskPaneRefreshCompletionMaterial material)
        {
            return new CreatedCaseDisplayCompletionDecision(
                TaskPaneRefreshCompletionDecisionStatus.ReadyForSession,
                string.Empty,
                material);
        }

        internal static CreatedCaseDisplayCompletionDecision Blocked(
            string blockedReason,
            TaskPaneRefreshCompletionMaterial material)
        {
            return new CreatedCaseDisplayCompletionDecision(
                TaskPaneRefreshCompletionDecisionStatus.Blocked,
                blockedReason,
                material);
        }
    }

    internal sealed class TaskPaneRefreshCompletionResult
    {
        private TaskPaneRefreshCompletionResult(
            TaskPaneRefreshCompletionResultStatus status,
            string resultReason,
            TaskPaneRefreshCompletionContext context,
            CreatedCaseDisplaySessionSnapshot sessionSnapshot)
        {
            Status = status;
            ResultReason = resultReason ?? string.Empty;
            Context = context;
            SessionSnapshot = sessionSnapshot;
            SessionId = sessionSnapshot == null ? string.Empty : sessionSnapshot.SessionId;
            WorkbookFullName = sessionSnapshot == null ? string.Empty : sessionSnapshot.WorkbookFullName;
        }

        internal TaskPaneRefreshCompletionResultStatus Status { get; }

        internal bool CanEmit
        {
            get
            {
                return Status == TaskPaneRefreshCompletionResultStatus.ReadyToEmit;
            }
        }

        internal string ResultReason { get; }

        internal TaskPaneRefreshCompletionContext Context { get; }

        internal CreatedCaseDisplaySessionSnapshot SessionSnapshot { get; }

        internal string SessionId { get; }

        internal string WorkbookFullName { get; }

        internal string Reason
        {
            get
            {
                return Context == null ? string.Empty : Context.Reason;
            }
        }

        internal TaskPaneRefreshAttemptResult AttemptResult
        {
            get
            {
                return Context == null ? null : Context.AttemptResult;
            }
        }

        internal string CompletionSource
        {
            get
            {
                return Context == null ? string.Empty : Context.CompletionSource;
            }
        }

        internal int? AttemptNumber
        {
            get
            {
                return Context == null ? null : Context.AttemptNumber;
            }
        }

        internal TaskPaneDisplayRequest DisplayRequest
        {
            get
            {
                return Context == null ? null : Context.DisplayRequest;
            }
        }

        internal Excel.Workbook Workbook
        {
            get
            {
                return Context == null ? null : Context.Workbook;
            }
        }

        internal Excel.Window Window
        {
            get
            {
                return Context == null ? null : Context.Window;
            }
        }

        internal static TaskPaneRefreshCompletionResult Blocked(
            TaskPaneRefreshCompletionContext context,
            string resultReason)
        {
            return new TaskPaneRefreshCompletionResult(
                TaskPaneRefreshCompletionResultStatus.Blocked,
                resultReason,
                context,
                sessionSnapshot: null);
        }

        internal static TaskPaneRefreshCompletionResult SessionMissing(
            TaskPaneRefreshCompletionContext context,
            string resultReason)
        {
            return new TaskPaneRefreshCompletionResult(
                TaskPaneRefreshCompletionResultStatus.SessionMissing,
                resultReason,
                context,
                sessionSnapshot: null);
        }

        internal static TaskPaneRefreshCompletionResult ReadyToEmit(
            TaskPaneRefreshCompletionContext context,
            CreatedCaseDisplaySessionSnapshot sessionSnapshot)
        {
            return new TaskPaneRefreshCompletionResult(
                TaskPaneRefreshCompletionResultStatus.ReadyToEmit,
                "readyToEmit",
                context,
                sessionSnapshot);
        }
    }
}
