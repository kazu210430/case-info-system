using System.Globalization;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneForegroundGuaranteeTraceBuilder
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private const string OrchestrationSource = "TaskPaneRefreshOrchestrationService";
        private const string ForegroundGuaranteeOutcomeSource =
            "TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome";

        internal TaskPaneForegroundGuaranteeTracePayload BuildDecisionTrace(
            TaskPaneForegroundGuaranteeDecisionTraceInput input)
        {
            TaskPaneRefreshAttemptResult attemptResult = input.Decision == null ? null : input.Decision.AttemptResult;
            ForegroundGuaranteeOutcome outcome = input.Decision == null ? null : input.Decision.Outcome;
            string details =
                "reason=" + (input.Reason ?? string.Empty)
                + ",foregroundRecoveryStarted=" + (input.Decision != null && input.Decision.ForegroundRecoveryStarted).ToString()
                + ",foregroundSkipReason=" + (input.Decision == null ? string.Empty : input.Decision.ForegroundSkipReason)
                + ",foregroundOutcomeStatus=" + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString());
            string message =
                KernelFlickerTracePrefix
                + " source=" + OrchestrationSource
                + " action=foreground-recovery-decision reason="
                + (input.Reason ?? string.Empty)
                + ", context="
                + input.FormattedContext
                + ", refreshSucceeded="
                + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ", resolvedWindowPresent="
                + (input.Decision != null && input.Decision.ResolvedWindow != null).ToString()
                + ", recoveryServicePresent="
                + (input.Decision != null && input.Decision.RecoveryServicePresent).ToString()
                + ", foregroundRecoveryStarted="
                + (input.Decision != null && input.Decision.ForegroundRecoveryStarted).ToString()
                + ", foregroundRecoverySkipped="
                + (input.Decision == null || !input.Decision.ForegroundRecoveryStarted).ToString()
                + ", foregroundSkipReason="
                + (input.Decision == null ? string.Empty : input.Decision.ForegroundSkipReason)
                + ", foregroundOutcomeStatus="
                + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString())
                + ", foregroundOutcomeDisplayCompletable="
                + (outcome != null && outcome.IsDisplayCompletable).ToString()
                + ", elapsedMs="
                + input.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + input.CorrelationFields;
            return new TaskPaneForegroundGuaranteeTracePayload(
                message,
                "foreground-recovery-decision",
                ForegroundGuaranteeOutcomeSource,
                details);
        }

        internal TaskPaneForegroundGuaranteeTracePayload BuildStartedTrace(
            TaskPaneForegroundGuaranteeStartedTraceInput input)
        {
            string details = "reason=" + (input.Reason ?? string.Empty);
            string message =
                KernelFlickerTracePrefix
                + " source=" + OrchestrationSource
                + " action=final-foreground-guarantee-start reason="
                + (input.Reason ?? string.Empty)
                + ", context="
                + input.FormattedContext
                + ", elapsedMs="
                + input.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + input.CorrelationFields;
            return new TaskPaneForegroundGuaranteeTracePayload(
                message,
                "final-foreground-guarantee-started",
                ForegroundGuaranteeOutcomeSource,
                details);
        }

        internal TaskPaneForegroundGuaranteeTracePayload BuildCompletedTrace(
            TaskPaneForegroundGuaranteeCompletedTraceInput input)
        {
            bool recovered = input.ExecutionResult != null && input.ExecutionResult.Recovered;
            string details =
                "reason=" + (input.Reason ?? string.Empty)
                + ",recovered=" + recovered.ToString()
                + ",foregroundOutcomeStatus="
                + (recovered
                    ? ForegroundGuaranteeOutcomeStatus.RequiredSucceeded.ToString()
                    : ForegroundGuaranteeOutcomeStatus.RequiredDegraded.ToString());
            string message =
                KernelFlickerTracePrefix
                + " source=" + OrchestrationSource
                + " action=final-foreground-guarantee-end reason="
                + (input.Reason ?? string.Empty)
                + ", context="
                + input.FormattedContext
                + ", recovered="
                + recovered.ToString()
                + ", elapsedMs="
                + input.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                + input.CorrelationFields;
            return new TaskPaneForegroundGuaranteeTracePayload(
                message,
                "final-foreground-guarantee-completed",
                ForegroundGuaranteeOutcomeSource,
                details);
        }
    }

    internal sealed class TaskPaneForegroundGuaranteeDecisionTraceInput
    {
        internal TaskPaneForegroundGuaranteeDecisionTraceInput(
            string reason,
            TaskPaneRefreshForegroundGuaranteeDecision decision,
            string formattedContext,
            long elapsedMilliseconds,
            string correlationFields)
        {
            Reason = reason;
            Decision = decision;
            FormattedContext = formattedContext ?? string.Empty;
            ElapsedMilliseconds = elapsedMilliseconds;
            CorrelationFields = correlationFields ?? string.Empty;
        }

        internal string Reason { get; }

        internal TaskPaneRefreshForegroundGuaranteeDecision Decision { get; }

        internal string FormattedContext { get; }

        internal long ElapsedMilliseconds { get; }

        internal string CorrelationFields { get; }
    }

    internal sealed class TaskPaneForegroundGuaranteeStartedTraceInput
    {
        internal TaskPaneForegroundGuaranteeStartedTraceInput(
            string reason,
            string formattedContext,
            long elapsedMilliseconds,
            string correlationFields)
        {
            Reason = reason;
            FormattedContext = formattedContext ?? string.Empty;
            ElapsedMilliseconds = elapsedMilliseconds;
            CorrelationFields = correlationFields ?? string.Empty;
        }

        internal string Reason { get; }

        internal string FormattedContext { get; }

        internal long ElapsedMilliseconds { get; }

        internal string CorrelationFields { get; }
    }

    internal sealed class TaskPaneForegroundGuaranteeCompletedTraceInput
    {
        internal TaskPaneForegroundGuaranteeCompletedTraceInput(
            string reason,
            ForegroundGuaranteeExecutionResult executionResult,
            string formattedContext,
            long elapsedMilliseconds,
            string correlationFields)
        {
            Reason = reason;
            ExecutionResult = executionResult;
            FormattedContext = formattedContext ?? string.Empty;
            ElapsedMilliseconds = elapsedMilliseconds;
            CorrelationFields = correlationFields ?? string.Empty;
        }

        internal string Reason { get; }

        internal ForegroundGuaranteeExecutionResult ExecutionResult { get; }

        internal string FormattedContext { get; }

        internal long ElapsedMilliseconds { get; }

        internal string CorrelationFields { get; }
    }

    internal sealed class TaskPaneForegroundGuaranteeTracePayload
    {
        internal TaskPaneForegroundGuaranteeTracePayload(
            string kernelTraceMessage,
            string observationAction,
            string observationSource,
            string details)
        {
            KernelTraceMessage = kernelTraceMessage ?? string.Empty;
            ObservationAction = observationAction ?? string.Empty;
            ObservationSource = observationSource ?? string.Empty;
            Details = details ?? string.Empty;
        }

        internal string KernelTraceMessage { get; }

        internal string ObservationAction { get; }

        internal string ObservationSource { get; }

        internal string Details { get; }
    }
}
