using System.Collections.Generic;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public sealed class WindowActivateDownstreamObservationTests
    {
        [Fact]
        public void LogStart_WhenWindowActivateRequest_EmitsObservationWithoutCompletionOwnership()
        {
            WindowActivateTaskPaneTriggerFacts facts = CreateFacts(out Excel.Workbook workbook, out Excel.Window window);
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate(facts);
            List<string> messages = new List<string>();
            WindowActivateDownstreamObservation observation = CreateObservation(messages);

            observation.LogStart(request, request.ToReasonString(), workbook, window, refreshAttemptId: 7);

            string message = Assert.Single(messages);
            Assert.Contains("source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-start", message);
            Assert.Contains("refreshAttemptId=7", message);
            Assert.Contains("windowActivateDispatchStatus=Dispatched", message);
            Assert.Contains("activationAttempt=NotAttempted", message);
            Assert.Contains("downstreamRecoveryDelegated=False", message);
            Assert.Contains("displayCompletionOutcome=False", message);
            Assert.Contains("recoveryOwner=False", message);
            Assert.Contains("foregroundGuaranteeOwner=False", message);
            Assert.Contains("hiddenExcelOwner=False", message);
            Assert.Contains("workbook=workbook-descriptor", message);
            Assert.Contains("inputWindow=window-descriptor", message);
            Assert.Contains("activeState=active-state", message);
            Assert.Contains("displayRequestSource=WindowActivate", message);
            Assert.Contains("windowActivateTriggerRole=TaskPaneDisplayRefreshTrigger", message);
            Assert.Contains("windowActivateCaptureOwner=WindowActivateDownstreamObservationTests", message);
            Assert.Contains("windowActivateWindowHwnd=123", message);
        }

        [Fact]
        public void LogOutcome_WhenDownstreamRecoveryWasDelegated_RemainsObservationNotCompletion()
        {
            WindowActivateTaskPaneTriggerFacts facts = CreateFacts(out _, out _);
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate(facts);
            TaskPaneRefreshAttemptResult attemptResult = TaskPaneRefreshAttemptResult.Failed(
                preContextRecoveryAttempted: true,
                preContextRecoverySucceeded: false);
            List<string> messages = new List<string>();
            WindowActivateDownstreamObservation observation = CreateObservation(messages);

            observation.LogOutcome(
                request,
                request.ToReasonString(),
                attemptResult,
                Stopwatch.StartNew(),
                refreshAttemptId: 8,
                completionSource: "refresh");

            string message = Assert.Single(messages);
            Assert.Contains("source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-outcome", message);
            Assert.Contains("refreshAttemptId=8", message);
            Assert.Contains("completionSource=refresh", message);
            Assert.Contains("windowActivateDispatchStatus=Dispatched", message);
            Assert.Contains("activationAttempt=Delegated", message);
            Assert.Contains("downstreamRecoveryDelegated=True", message);
            Assert.Contains("displayCompletionOutcome=False", message);
            Assert.Contains("recoveryOwner=False", message);
            Assert.Contains("foregroundGuaranteeOwner=False", message);
            Assert.Contains("hiddenExcelOwner=False", message);
            Assert.Contains("refreshSucceeded=False", message);
            Assert.Contains("preContextFullRecoveryAttempted=True", message);
            Assert.Contains("displayRequestSource=WindowActivate", message);
        }

        [Fact]
        public void LogOutcome_WhenRefreshFactsAreDisplayCompletable_RemainsObservationNotCompletion()
        {
            WindowActivateTaskPaneTriggerFacts facts = CreateFacts(out _, out _);
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForWindowActivate(facts);
            TaskPaneRefreshAttemptResult attemptResult = TaskPaneRefreshAttemptResult
                .VisibleAlreadySatisfied()
                .WithVisibilityRecoveryOutcome(VisibilityRecoveryOutcome.Completed(
                    "refreshedShown",
                    VisibilityRecoveryTargetKind.ExplicitWorkbookWindow,
                    PaneVisibleSource.RefreshedShown,
                    WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible,
                    fullRecoveryAttempted: false,
                    fullRecoverySucceeded: null))
                .WithForegroundGuaranteeOutcome(ForegroundGuaranteeOutcome.RequiredSucceeded(
                    ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                    "foregroundRecoverySucceeded"));
            List<string> messages = new List<string>();
            WindowActivateDownstreamObservation observation = CreateObservation(messages);

            observation.LogOutcome(
                request,
                request.ToReasonString(),
                attemptResult,
                Stopwatch.StartNew(),
                refreshAttemptId: 9,
                completionSource: "refresh");

            string message = Assert.Single(messages);
            Assert.Contains("source=TaskPaneRefreshOrchestrationService action=window-activate-display-refresh-trigger-outcome", message);
            Assert.Contains("refreshAttemptId=9", message);
            Assert.Contains("displayCompletionOutcome=False", message);
            Assert.Contains("recoveryOwner=False", message);
            Assert.Contains("foregroundGuaranteeOwner=False", message);
            Assert.Contains("hiddenExcelOwner=False", message);
            Assert.Contains("refreshSucceeded=True", message);
            Assert.Contains("paneVisible=True", message);
            Assert.Contains("visibilityRecoveryStatus=Completed", message);
            Assert.Contains("foregroundGuaranteeStatus=RequiredSucceeded", message);
            Assert.DoesNotContain("case-display-completed", message);
        }

        [Fact]
        public void LogStartAndOutcome_WhenRequestIsNotWindowActivate_DoNotEmitDownstreamObservation()
        {
            TaskPaneDisplayRequest request = TaskPaneDisplayRequest.ForPostActionRefresh("doc");
            List<string> messages = new List<string>();
            WindowActivateDownstreamObservation observation = CreateObservation(messages);

            observation.LogStart(request, request.ToReasonString(), workbook: null, window: null, refreshAttemptId: 1);
            observation.LogOutcome(
                request,
                request.ToReasonString(),
                TaskPaneRefreshAttemptResult.Succeeded(),
                Stopwatch.StartNew(),
                refreshAttemptId: 1,
                completionSource: "refresh");

            Assert.Empty(messages);
            Assert.Contains(
                "displayRequestSource=PostActionRefresh",
                WindowActivateDownstreamObservation.FormatDisplayRequestTraceFields(request));
        }

        private static WindowActivateDownstreamObservation CreateObservation(List<string> messages)
        {
            return new WindowActivateDownstreamObservation(
                new Logger(messages.Add),
                _ => "workbook-descriptor",
                _ => "window-descriptor",
                () => "active-state");
        }

        private static WindowActivateTaskPaneTriggerFacts CreateFacts(out Excel.Workbook workbook, out Excel.Window window)
        {
            var application = new Excel.Application();
            workbook = new Excel.Workbook
            {
                Application = application,
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx"
            };
            application.Workbooks.Add(workbook);
            window = workbook.Windows[1];
            window.Hwnd = 123;
            window.Caption = "case.xlsx";
            return new WindowActivateTaskPaneTriggerFacts(
                workbook,
                window,
                "full=\"C:\\cases\\case.xlsx\",name=\"case.xlsx\"",
                "hwnd=\"123\",caption=\"case.xlsx\"",
                "activeWorkbook=full=\"C:\\cases\\case.xlsx\",activeWindow=hwnd=\"123\"",
                workbook.FullName,
                window.Hwnd.ToString(),
                "WindowActivateDownstreamObservationTests");
        }
    }
}
