using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class TaskPaneRefreshObservationDecisionServiceTests
    {
        [Fact]
        public void CompleteNormalizedOutcomeChain_ConnectsObservationFactsToOutcomeChain()
        {
            var service = new TaskPaneRefreshObservationDecisionService();
            TaskPaneRefreshAttemptResult rawAttempt = TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied();

            TaskPaneRefreshObservationDecision decision = service.CompleteNormalizedOutcomeChain(
                new TaskPaneRefreshObservationDecisionInput(
                    "ready-show",
                    workbook: null,
                    inputWindow: new Excel.Window { Hwnd = 101 },
                    attemptResult: rawAttempt,
                    completionSource: "ready-show-attempt",
                    attemptNumber: 2,
                    workbookWindowEnsureFacts: null));

            Assert.True(decision.AttemptResult.IsRefreshSucceeded);
            Assert.Equal(VisibilityRecoveryOutcomeStatus.Skipped, decision.Visibility.Outcome.Status);
            Assert.True(decision.Visibility.Outcome.IsDisplayCompletable);
            Assert.Equal(RefreshSourceSelectionOutcomeStatus.NotReached, decision.RefreshSource.Outcome.Status);
            Assert.Equal("refresh-source-not-reached", decision.RefreshSource.StatusAction);
            Assert.Equal(RebuildFallbackOutcomeStatus.Skipped, decision.RebuildFallback.Outcome.Status);
            Assert.Equal("rebuild-fallback-skipped", decision.RebuildFallback.StatusAction);
            Assert.Contains("completionSource=ready-show-attempt", decision.Visibility.Details);
            Assert.Contains("attempt=2", decision.Visibility.Details);
        }

        [Fact]
        public void DecideForegroundGuarantee_WhenWindowAndRecoveryServiceAreAvailable_ReturnsExecuteDecisionOnly()
        {
            var service = new TaskPaneRefreshObservationDecisionService();
            Excel.Workbook workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx"
            };
            Excel.Window window = new Excel.Window { Hwnd = 202 };
            var context = new WorkbookContext(
                workbook,
                window,
                WorkbookRole.Case,
                @"C:\cases",
                @"C:\cases\case.xlsx",
                "shHOME");
            TaskPaneRefreshAttemptResult attempt = TaskPaneRefreshAttemptResult.RefreshCompletedPendingForeground(
                context,
                workbook,
                window,
                isForegroundRecoveryServiceAvailable: true,
                completionBasis: "refreshCompleted",
                paneVisibleSource: PaneVisibleSource.RefreshedShown,
                snapshotBuildResult: null,
                preContextRecoveryAttempted: false,
                preContextRecoverySucceeded: null);

            TaskPaneRefreshForegroundGuaranteeDecision decision = service.DecideForegroundGuarantee(
                attempt,
                inputWindow: window);

            Assert.True(decision.ShouldExecuteForegroundGuarantee);
            Assert.True(decision.ForegroundRecoveryStarted);
            Assert.Equal(ForegroundGuaranteeOutcomeStatus.Unknown, decision.Outcome.Status);
            Assert.Equal(ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow, decision.TargetKind);
            Assert.Same(context, decision.Context);
            Assert.Same(window, decision.ResolvedWindow);
            Assert.Contains(
                "foregroundRecoveryStarted=True",
                service.BuildForegroundRecoveryDecisionDetails("ready-show", decision));
        }

        [Fact]
        public void DecideForegroundGuarantee_WhenRefreshFailed_ReturnsNonExecutingNotRequiredDecision()
        {
            var service = new TaskPaneRefreshObservationDecisionService();

            TaskPaneRefreshForegroundGuaranteeDecision decision = service.DecideForegroundGuarantee(
                TaskPaneRefreshAttemptResult.Failed(),
                inputWindow: null);

            Assert.False(decision.ShouldExecuteForegroundGuarantee);
            Assert.False(decision.ForegroundRecoveryStarted);
            Assert.Equal(ForegroundGuaranteeOutcomeStatus.NotRequired, decision.Outcome.Status);
            Assert.Equal("refreshSucceeded=false", decision.ForegroundSkipReason);
            Assert.Equal(decision.Outcome, decision.AttemptResult.ForegroundGuaranteeOutcome);
        }

        [Fact]
        public void ClassifyRequiredForegroundExecutionOutcome_DegradedRemainsDisplayCompletable()
        {
            var service = new TaskPaneRefreshObservationDecisionService();

            ForegroundGuaranteeOutcome outcome = service.ClassifyRequiredForegroundExecutionOutcome(
                ForegroundGuaranteeTargetKind.ExplicitWorkbookWindow,
                new ForegroundGuaranteeExecutionResult(executionAttempted: true, recovered: false, elapsedMilliseconds: 1));

            Assert.Equal(ForegroundGuaranteeOutcomeStatus.RequiredDegraded, outcome.Status);
            Assert.True(outcome.IsTerminal);
            Assert.True(outcome.IsDisplayCompletable);
            Assert.False(outcome.RecoverySucceeded.Value);
        }

        [Fact]
        public void Source_DoesNotOwnTimerCallbackCompletionOrDisplayExecution()
        {
            string source = ReadAppSource("TaskPaneRefreshObservationDecisionService.cs");

            Assert.Contains("TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(", source);
            Assert.Contains("TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(", source);
            Assert.Contains("TaskPaneNormalizedOutcomeMapper.BuildRebuildFallbackOutcome(", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("TaskPaneReadyShowRetryScheduler", source);
            Assert.DoesNotContain("PendingPaneRefreshRetryService", source);
            Assert.DoesNotContain("WindowActivateDownstreamObservation", source);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("case-display-completed", source);
            Assert.DoesNotContain("TaskPaneRefreshCoordinator", source);
            Assert.DoesNotContain("TaskPaneManager", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
            Assert.DoesNotContain("BeginPostForegroundProtection", source);
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
