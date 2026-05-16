using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneForegroundGuaranteeTraceBuilderTests
    {
        [Fact]
        public void BuildDecisionTrace_BuildsKernelTraceAndObservationDetails()
        {
            var builder = new TaskPaneForegroundGuaranteeTraceBuilder();
            TaskPaneRefreshForegroundGuaranteeDecision decision =
                TaskPaneRefreshForegroundGuaranteeDecision.NoExecution(
                    TaskPaneRefreshAttemptResult.Succeeded(),
                    ForegroundGuaranteeOutcome.NotRequired("refreshCompleted=false"),
                    inputWindow: null,
                    foregroundSkipReason: "refreshCompleted=false");

            TaskPaneForegroundGuaranteeTracePayload payload = builder.BuildDecisionTrace(
                new TaskPaneForegroundGuaranteeDecisionTraceInput(
                    "ready-show",
                    decision,
                    "context=case",
                    elapsedMilliseconds: 15,
                    correlationFields: ",correlationId=abc"));

            Assert.Contains("action=foreground-recovery-decision", payload.KernelTraceMessage);
            Assert.Contains(", foregroundRecoverySkipped=True", payload.KernelTraceMessage);
            Assert.Contains(",correlationId=abc", payload.KernelTraceMessage);
            Assert.Equal("foreground-recovery-decision", payload.ObservationAction);
            Assert.Equal("TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome", payload.ObservationSource);
            Assert.Equal(
                "reason=ready-show,foregroundRecoveryStarted=False,foregroundSkipReason=refreshCompleted=false,foregroundOutcomeStatus=NotRequired",
                payload.Details);
        }

        [Fact]
        public void BuildCompletedTrace_MapsRecoveredToSucceededOtherwiseDegraded()
        {
            var builder = new TaskPaneForegroundGuaranteeTraceBuilder();

            TaskPaneForegroundGuaranteeTracePayload succeeded = builder.BuildCompletedTrace(
                new TaskPaneForegroundGuaranteeCompletedTraceInput(
                    "ready-show",
                    new ForegroundGuaranteeExecutionResult(executionAttempted: true, recovered: true, elapsedMilliseconds: 1),
                    "context=case",
                    elapsedMilliseconds: 20,
                    correlationFields: string.Empty));
            TaskPaneForegroundGuaranteeTracePayload degraded = builder.BuildCompletedTrace(
                new TaskPaneForegroundGuaranteeCompletedTraceInput(
                    "ready-show",
                    new ForegroundGuaranteeExecutionResult(executionAttempted: true, recovered: false, elapsedMilliseconds: 1),
                    "context=case",
                    elapsedMilliseconds: 20,
                    correlationFields: string.Empty));

            Assert.Contains("foregroundOutcomeStatus=RequiredSucceeded", succeeded.Details);
            Assert.Contains("foregroundOutcomeStatus=RequiredDegraded", degraded.Details);
            Assert.DoesNotContain("RequiredFailed", degraded.Details);
        }

        [Fact]
        public void Source_DoesNotOwnForegroundExecutionEmitOrSessionLifecycle()
        {
            string source = ReadAppSource("TaskPaneForegroundGuaranteeTraceBuilder.cs");

            Assert.Contains("action=foreground-recovery-decision", source);
            Assert.Contains("action=final-foreground-guarantee-start", source);
            Assert.Contains("action=final-foreground-guarantee-end", source);
            Assert.DoesNotContain("_logger", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
            Assert.DoesNotContain("BeginPostForegroundProtection", source);
            Assert.DoesNotContain("_createdCaseDisplaySessions", source);
            Assert.DoesNotContain("case-display-completed", source);
        }

        private static string ReadAppSource(string appFileName)
        {
            string repoRoot = FindRepositoryRoot();
            return File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", appFileName));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
