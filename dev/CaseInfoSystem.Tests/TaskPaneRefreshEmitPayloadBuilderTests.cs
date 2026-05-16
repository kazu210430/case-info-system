using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRefreshEmitPayloadBuilderTests
    {
        [Fact]
        public void BuildCaseDisplayCompleted_BuildsTraceAndObservationPayloadWithoutEmitting()
        {
            var builder = new TaskPaneRefreshEmitPayloadBuilder();
            TaskPaneRefreshAttemptResult attempt = TaskPaneRefreshAttemptResult.Succeeded()
                .WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome.Completed(
                    "visible",
                    VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                    PaneVisibleSource.RefreshedShown,
                    workbookWindowEnsureStatus: null,
                    fullRecoveryAttempted: false,
                    fullRecoverySucceeded: null));

            CaseDisplayCompletedPayload payload = builder.BuildCaseDisplayCompleted(
                new CaseDisplayCompletedPayloadInput(
                    "ready-show",
                    "CDS-0001",
                    @"C:\cases\case.xlsx",
                    attempt,
                    "ready-show-attempt",
                    2,
                    displayRequest: null,
                    formattedWorkbook: "full=\"case.xlsx\"",
                    formattedWindow: "hwnd=\"100\""));

            Assert.Contains("action=case-display-completed sessionId=CDS-0001", payload.KernelTraceMessage);
            Assert.Equal("case-display-completed", payload.ObservationAction);
            Assert.Equal("TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession", payload.ObservationSource);
            Assert.Equal(@"C:\cases\case.xlsx", payload.WorkbookFullName);
            Assert.Contains("reason=ready-show,sessionId=CDS-0001,completionSource=ready-show-attempt", payload.Details);
            Assert.Contains("foregroundGuaranteeStatus=NotRequired", payload.Details);
            Assert.Contains("attempt=2", payload.Details);
        }

        [Fact]
        public void BuildCaseDisplayCompleted_WithEmitContext_PreservesExistingPayloadFields()
        {
            var payloadBuilder = new TaskPaneRefreshEmitPayloadBuilder();
            var contextBuilder = new TaskPaneRefreshEmitContextBuilder(
                workbook => "full=\"case-context.xlsx\"",
                window => "hwnd=\"200\"");
            var session = new CreatedCaseDisplaySession(
                "CDS-0002",
                @"C:\cases\case-context.xlsx",
                "ready-show");
            TaskPaneRefreshAttemptResult attempt = TaskPaneRefreshAttemptResult.Succeeded()
                .WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome.Completed(
                    "visible",
                    VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                    PaneVisibleSource.RefreshedShown,
                    workbookWindowEnsureStatus: null,
                    fullRecoveryAttempted: false,
                    fullRecoverySucceeded: null));

            TaskPaneRefreshEmitContext context = contextBuilder.Build(
                TaskPaneRefreshEmitContextInput.ForCompletedSession(
                    "ready-show",
                    session,
                    attempt,
                    "ready-show-attempt",
                    2,
                    displayRequest: null,
                    workbook: null,
                    window: null));
            CaseDisplayCompletedPayload payload = payloadBuilder.BuildCaseDisplayCompleted(context);

            Assert.Equal("CDS-0002", context.SessionId);
            Assert.Equal(@"C:\cases\case-context.xlsx", context.WorkbookFullName);
            Assert.Equal("full=\"case-context.xlsx\"", context.FormattedWorkbook);
            Assert.Equal("hwnd=\"200\"", context.FormattedWindow);
            Assert.Contains("action=case-display-completed sessionId=CDS-0002", payload.KernelTraceMessage);
            Assert.Contains("workbook=full=\"case-context.xlsx\", window=hwnd=\"200\"", payload.KernelTraceMessage);
            Assert.Equal(@"C:\cases\case-context.xlsx", payload.WorkbookFullName);
            Assert.Contains("reason=ready-show,sessionId=CDS-0002,completionSource=ready-show-attempt", payload.Details);
            Assert.Contains("attempt=2", payload.Details);
        }

        [Fact]
        public void Source_DoesNotOwnEmitSessionMutationOrCompletionDecision()
        {
            string source = ReadAppSource("TaskPaneRefreshEmitPayloadBuilder.cs");

            Assert.Contains("action=case-display-completed", source);
            Assert.DoesNotContain("_logger", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("_createdCaseDisplaySessions", source);
            Assert.DoesNotContain("IsCompleted", source);
            Assert.DoesNotContain("CanComplete", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
        }

        [Fact]
        public void EmitContext_SourceOnlyBuildsMaterialWithoutOwningEmitOrLifecycle()
        {
            string source = ReadAppSource("TaskPaneRefreshEmitContext.cs");

            Assert.Contains("TaskPaneRefreshEmitContextBuilder", source);
            Assert.Contains("session == null ? null : session.ToSnapshot()", source);
            Assert.DoesNotContain("_logger", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("_createdCaseDisplaySessions", source);
            Assert.DoesNotContain("IsCompleted = true", source);
            Assert.DoesNotContain("Remove(", source);
            Assert.DoesNotContain("DecideCreatedCaseDisplayCompletion", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
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
