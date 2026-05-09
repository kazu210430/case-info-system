using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneDisplayRetryCoordinatorTests
    {
        [Fact]
        public void ShowWhenReady_RetriesThroughSchedulerUntilAttemptSucceeds()
        {
            var attempts = new List<int>();
            var scheduledAttempts = new List<int>();
            bool shown = false;
            bool fallbackScheduled = false;
            var coordinator = new TaskPaneDisplayRetryCoordinator(maxAttempts: 3);

            coordinator.ShowWhenReady(
                workbook: null,
                reason: "test",
                tryShowOnce: (_, __, attemptNumber) =>
                {
                    attempts.Add(attemptNumber);
                    return attemptNumber == 3;
                },
                scheduleRetry: (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                onShown: () => shown = true,
                scheduleFallback: (_, __) => fallbackScheduled = true);

            Assert.Equal(new[] { 1, 2, 3 }, attempts);
            Assert.Equal(new[] { 2, 3 }, scheduledAttempts);
            Assert.True(shown);
            Assert.False(fallbackScheduled);
        }

        [Fact]
        public void ShowWhenReady_WhenReadyShowAttemptsExhausted_SchedulesFallbackWithoutAttempt3()
        {
            var attempts = new List<int>();
            var scheduledAttempts = new List<int>();
            bool shown = false;
            bool fallbackScheduled = false;
            string fallbackReason = null;
            var coordinator = new TaskPaneDisplayRetryCoordinator(WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts);

            coordinator.ShowWhenReady(
                workbook: null,
                reason: "fallback",
                tryShowOnce: (_, __, attemptNumber) =>
                {
                    attempts.Add(attemptNumber);
                    return false;
                },
                scheduleRetry: (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                onShown: () => shown = true,
                scheduleFallback: (_, reason) =>
                {
                    fallbackScheduled = true;
                    fallbackReason = reason;
                });

            Assert.Equal(new[] { 1, 2 }, attempts);
            Assert.Equal(new[] { 2 }, scheduledAttempts);
            Assert.DoesNotContain(3, attempts);
            Assert.DoesNotContain(3, scheduledAttempts);
            Assert.False(shown);
            Assert.True(fallbackScheduled);
            Assert.Equal("fallback", fallbackReason);
        }

        [Fact]
        public void ReadyShowRetryContract_UsesTwoAttemptsSeparateFromPendingRetry()
        {
            Assert.Equal(2, WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts);
            Assert.Equal(80, WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowRetryDelayMs);

            string repoRoot = FindRepositoryRoot();
            string compositionSource = File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "AddInCompositionRoot.cs"));
            string orchestrationSource = File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", "TaskPaneRefreshOrchestrationService.cs"));
            string thisAddInSource = File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "ThisAddIn.cs"));

            Assert.Contains("new TaskPaneDisplayRetryCoordinator(ReadyShowRetryMaxAttempts)", compositionSource);
            Assert.DoesNotContain("new TaskPaneDisplayRetryCoordinator(_pendingPaneRefreshMaxAttempts)", compositionSource);
            Assert.Contains("new TaskPaneReadyShowRetryScheduler", orchestrationSource);
            Assert.DoesNotContain("private void ScheduleTaskPaneReadyRetry", orchestrationSource);
            Assert.DoesNotContain("TaskPaneRefreshOrchestrationService.PendingPaneRefreshMaxAttempts", thisAddInSource);
            Assert.Contains("internal const int PendingPaneRefreshIntervalMs = 400;", orchestrationSource);
            Assert.Contains("internal const int PendingPaneRefreshMaxAttempts = 3;", orchestrationSource);
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

    public sealed class DisplayConvergenceCompletionBoundarySourceTests
    {
        [Fact]
        public void CaseDisplayCompleted_EmitOwnerAndOneTimeGateStayInOrchestrationOnly()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");

            Assert.Contains("TryCompleteCreatedCaseDisplaySession", orchestrationSource);
            Assert.Contains("action=case-display-completed", orchestrationSource);
            Assert.Contains("\"case-display-completed\"", orchestrationSource);
            Assert.Contains("\"TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession\"", orchestrationSource);
            Assert.Contains("NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);", orchestrationSource);
            AssertContainsInOrder(
                orchestrationSource,
                "bool shouldEmit = false;",
                "if (!resolvedSession.IsCompleted)",
                "resolvedSession.IsCompleted = true;",
                "_createdCaseDisplaySessions.Remove(resolvedSession.WorkbookFullName);",
                "shouldEmit = true;",
                "if (!shouldEmit)",
                "return;",
                "string details =");

            AssertDoesNotOwnCompletion("WorkbookTaskPaneReadyShowAttemptWorker.cs");
            AssertDoesNotOwnCompletion("PendingPaneRefreshRetryService.cs");
            AssertDoesNotOwnCompletion("TaskPaneDisplayRetryCoordinator.cs");
            AssertDoesNotOwnCompletion("TaskPaneRefreshCoordinator.cs");
            AssertDoesNotOwnCompletion("WindowActivatePaneHandlingService.cs");
            AssertDoesNotOwnCompletion("WindowActivateDownstreamObservation.cs");
        }

        [Fact]
        public void CompletionHardGateDecisionContract_RequiresVisibilityAndForegroundDisplayCompletableFacts()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");

            Assert.Contains("IsCreatedCaseDisplayReason(reason)", orchestrationSource);
            Assert.Contains("attemptResult == null", orchestrationSource);
            Assert.Contains("!attemptResult.IsRefreshSucceeded", orchestrationSource);
            Assert.Contains("!attemptResult.IsPaneVisible", orchestrationSource);
            Assert.Contains("attemptResult.VisibilityRecoveryOutcome == null", orchestrationSource);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsTerminal", orchestrationSource);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable", orchestrationSource);
            Assert.Contains("!attemptResult.IsForegroundGuaranteeTerminal", orchestrationSource);
            Assert.Contains("attemptResult.ForegroundGuaranteeOutcome == null", orchestrationSource);
            Assert.Contains("!attemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable", orchestrationSource);
        }

        [Fact]
        public void CaseDisplayCompletedDetailsPayload_PreservesFieldSetAndOrder()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string payloadHelper = Slice(
                orchestrationSource,
                "private static string BuildCaseDisplayCompletedDetailsPayload",
                "private static CreatedCaseDisplayCompletionDecision");

            AssertContainsInOrder(
                orchestrationSource,
                "if (!completionDecision.CanComplete)",
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);",
                "bool shouldEmit = false;",
                "if (!shouldEmit)",
                "string details = BuildCaseDisplayCompletedDetailsPayload(",
                "_logger?.Info(",
                "\"case-display-completed\"",
                "details);",
                "NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);");
            AssertContainsInOrder(
                payloadHelper,
                "string details =",
                "\"reason=\" + (reason ?? string.Empty)",
                "\",sessionId=\" + resolvedSession.SessionId",
                "\",completionSource=\" + (completionSource ?? string.Empty)",
                "\",completion=\" + attemptResult.CompletionBasis",
                "\",paneVisible=\" + attemptResult.IsPaneVisible.ToString()",
                "\",visibilityRecoveryStatus=\" + attemptResult.VisibilityRecoveryOutcome.Status.ToString()",
                "\",visibilityRecoveryDisplayCompletable=\" + attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable.ToString()",
                "\",visibilityRecoveryPaneVisible=\" + attemptResult.VisibilityRecoveryOutcome.IsPaneVisible.ToString()",
                "\",visibilityRecoveryTargetKind=\" + attemptResult.VisibilityRecoveryOutcome.TargetKind.ToString()",
                "\",visibilityPaneVisibleSource=\" + attemptResult.VisibilityRecoveryOutcome.PaneVisibleSource.ToString()",
                "\",visibilityRecoveryReason=\" + attemptResult.VisibilityRecoveryOutcome.Reason",
                "\",visibilityRecoveryDegradedReason=\" + attemptResult.VisibilityRecoveryOutcome.DegradedReason",
                "\",refreshSourceStatus=\" + attemptResult.RefreshSourceSelectionOutcome.Status.ToString()",
                "\",refreshSourceSelectedSource=\" + attemptResult.RefreshSourceSelectionOutcome.SelectedSource.ToString()",
                "\",refreshSourceSelectionReason=\" + attemptResult.RefreshSourceSelectionOutcome.SelectionReason",
                "\",refreshSourceFallbackReasons=\" + attemptResult.RefreshSourceSelectionOutcome.FallbackReasons",
                "\",refreshSourceCacheFallback=\" + attemptResult.RefreshSourceSelectionOutcome.IsCacheFallback.ToString()",
                "\",refreshSourceRebuildRequired=\" + attemptResult.RefreshSourceSelectionOutcome.IsRebuildRequired.ToString()",
                "\",refreshSourceCanContinue=\" + attemptResult.RefreshSourceSelectionOutcome.CanContinueRefresh.ToString()",
                "\",refreshSourceFailureReason=\" + attemptResult.RefreshSourceSelectionOutcome.FailureReason",
                "\",refreshSourceDegradedReason=\" + attemptResult.RefreshSourceSelectionOutcome.DegradedReason",
                "\",rebuildFallbackStatus=\" + attemptResult.RebuildFallbackOutcome.Status.ToString()",
                "\",rebuildFallbackRequired=\" + attemptResult.RebuildFallbackOutcome.IsRequired.ToString()",
                "\",rebuildFallbackCanContinue=\" + attemptResult.RebuildFallbackOutcome.CanContinueRefresh.ToString()",
                "\",rebuildFallbackSnapshotSource=\" + attemptResult.RebuildFallbackOutcome.SnapshotSource.ToString()",
                "\",rebuildFallbackReasons=\" + attemptResult.RebuildFallbackOutcome.FallbackReasons",
                "\",rebuildFallbackFailureReason=\" + attemptResult.RebuildFallbackOutcome.FailureReason",
                "\",rebuildFallbackDegradedReason=\" + attemptResult.RebuildFallbackOutcome.DegradedReason",
                "\",refreshCompleted=\" + attemptResult.IsRefreshCompleted.ToString()",
                "\",foregroundGuaranteeTerminal=\" + attemptResult.IsForegroundGuaranteeTerminal.ToString()",
                "\",foregroundGuaranteeRequired=\" + attemptResult.WasForegroundGuaranteeRequired.ToString()",
                "\",foregroundGuaranteeStatus=\" + attemptResult.ForegroundGuaranteeOutcome.Status.ToString()",
                "\",foregroundGuaranteeDisplayCompletable=\" + attemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable.ToString()",
                "\",foregroundGuaranteeExecutionAttempted=\" + attemptResult.ForegroundGuaranteeOutcome.WasExecutionAttempted.ToString()",
                "\",foregroundGuaranteeTargetKind=\" + attemptResult.ForegroundGuaranteeOutcome.TargetKind.ToString()",
                "\",foregroundRecoverySucceeded=\"",
                "\",foregroundOutcomeReason=\" + attemptResult.ForegroundGuaranteeOutcome.Reason",
                "WindowActivateDownstreamObservation.FormatDisplayRequestTraceFields(displayRequest)",
                "details += \",attempt=\" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);",
                "return details;");
            Assert.DoesNotContain("_logger", payloadHelper);
            Assert.DoesNotContain("NewCaseVisibilityObservation", payloadHelper);
            Assert.DoesNotContain("_createdCaseDisplaySessions", payloadHelper);
            Assert.DoesNotContain("IsCompleted", payloadHelper);
        }

        [Fact]
        public void NormalRefreshPath_KeepsNormalizedOutcomeChainBeforeForegroundWindowActivateAndCompletion()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string normalRefreshPath = Slice(
                orchestrationSource,
                "RefreshDispatchExecutionResult dispatchExecutionResult = RefreshDispatchShell.Dispatch(",
                "return attemptResult;");

            AssertNormalizedOutcomeChainBefore(
                normalRefreshPath,
                "CompleteForegroundGuaranteeOutcome(",
                orchestrationSource);
            AssertContainsInOrder(
                normalRefreshPath,
                "CompleteForegroundGuaranteeOutcome(",
                "_windowActivateDownstreamObservation.LogOutcome(",
                "TryCompleteCreatedCaseDisplaySession(");
        }

        [Fact]
        public void PreconditionSkipPath_KeepsNormalizedOutcomeChainBeforeWindowActivateAndReturnWithoutForegroundOrCompletion()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string preconditionSkipPath = Slice(
                orchestrationSource,
                "if (!preconditionDecision.CanRefresh)",
                "RefreshDispatchExecutionResult dispatchExecutionResult = RefreshDispatchShell.Dispatch(");

            AssertNormalizedOutcomeChainBefore(
                preconditionSkipPath,
                "_windowActivateDownstreamObservation.LogOutcome(",
                orchestrationSource);
            AssertContainsInOrder(
                preconditionSkipPath,
                "_windowActivateDownstreamObservation.LogOutcome(",
                "return skippedResult;");
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", preconditionSkipPath);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", preconditionSkipPath);
            Assert.DoesNotContain("case-display-completed", preconditionSkipPath);
            Assert.DoesNotContain("foreground-recovery-decision", preconditionSkipPath);
            Assert.DoesNotContain("final-foreground-guarantee", preconditionSkipPath);
        }

        [Fact]
        public void ReadyShowCallback_ReturnsFactsToConvergenceChainBeforeCompletionGate()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string readyShowHandoff = Slice(
                orchestrationSource,
                "CreatedCaseDisplaySession createdCaseDisplaySession = BeginCreatedCaseDisplaySession",
                "internal Excel.Window ResolveWorkbookPaneWindow");
            string callbackHandler = Slice(
                orchestrationSource,
                "private void HandleWorkbookTaskPaneShown",
                "private void TryCompleteCreatedCaseDisplaySession");
            string workerSource = ReadAppSource("WorkbookTaskPaneReadyShowAttemptWorker.cs");

            Assert.Contains(
                "outcome => HandleWorkbookTaskPaneShown(createdCaseDisplaySession, workbook, reason, outcome)",
                readyShowHandoff);
            AssertNormalizedOutcomeChainBefore(
                callbackHandler,
                "CompleteForegroundGuaranteeOutcome(",
                orchestrationSource);
            AssertContainsInOrder(
                callbackHandler,
                "CompleteForegroundGuaranteeOutcome(",
                "TryCompleteCreatedCaseDisplaySession(");
            Assert.Contains("() => onShown?.Invoke(shownOutcome)", workerSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", workerSource);
            Assert.DoesNotContain("case-display-completed", workerSource);
        }

        [Fact]
        public void NormalizedOutcomeChainMethods_DoNotOwnCompletionSessionOrOneTimeEmit()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string normalizedOutcomeChainSource =
                ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteVisibilityRecoveryOutcome")
                + ReadMethod(orchestrationSource, "private void LogVisibilityRecoveryOutcome")
                + ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteRefreshSourceSelectionOutcome")
                + ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionOutcome")
                + ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionTrace")
                + ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteRebuildFallbackOutcome")
                + ReadMethod(orchestrationSource, "private void LogRebuildFallbackOutcome")
                + ReadOptionalMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteNormalizedOutcomeChain");

            Assert.DoesNotContain("case-display-completed", normalizedOutcomeChainSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", normalizedOutcomeChainSource);
            Assert.DoesNotContain("ResolveCreatedCaseDisplaySession", normalizedOutcomeChainSource);
            Assert.DoesNotContain("_createdCaseDisplaySessions", normalizedOutcomeChainSource);
            Assert.DoesNotContain("IsCompleted", normalizedOutcomeChainSource);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", normalizedOutcomeChainSource);
            Assert.DoesNotContain("EvaluateCreatedCaseDisplayCompletionDecision", normalizedOutcomeChainSource);
        }

        [Fact]
        public void NormalizedOutcomeTraceActionsAndSources_RemainR10R11R12Specific()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string visibilityTrace = ReadMethod(orchestrationSource, "private void LogVisibilityRecoveryOutcome");
            string refreshSourceTrace = ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionOutcome")
                + ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionTrace");
            string rebuildFallbackTrace = ReadMethod(orchestrationSource, "private void LogRebuildFallbackOutcome");

            AssertContainsInOrder(
                visibilityTrace,
                "TaskPaneNormalizedOutcomeMapper.FormatVisibilityRecoveryDetails(",
                "action=visibility-recovery-decision",
                "\"visibility-recovery-decision\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome\"",
                "\"visibility-recovery-\" + outcome.Status.ToString().ToLowerInvariant()",
                "\"TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome\"");
            AssertContainsInOrder(
                refreshSourceTrace,
                "TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionDetails(",
                "TaskPaneNormalizedOutcomeMapper.FormatRefreshSourceSelectionAction(outcome)",
                "\"refresh-source-rebuild-required\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteRefreshSourceSelectionOutcome\"");
            AssertContainsInOrder(
                rebuildFallbackTrace,
                "TaskPaneNormalizedOutcomeMapper.FormatRebuildFallbackDetails(",
                "\"rebuild-fallback-required\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome\"",
                "string statusAction = \"rebuild-fallback-\" + outcome.Status.ToString().ToLowerInvariant()",
                "\"TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome\"");
        }

        [Fact]
        public void ForegroundOutcomeChain_DoesNotOwnCompletionSessionOrOneTimeEmit()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string foregroundOutcomeChainSource =
                ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteForegroundGuaranteeOutcome")
                + ReadMethod(orchestrationSource, "private ForegroundGuaranteeOutcome ExecuteForegroundGuaranteeAndBuildOutcome")
                + ReadMethod(orchestrationSource, "private void LogForegroundGuaranteeDecision")
                + ReadMethod(orchestrationSource, "private void LogFinalForegroundGuaranteeStarted")
                + ReadMethod(orchestrationSource, "private void LogFinalForegroundGuaranteeCompleted");

            Assert.DoesNotContain("case-display-completed", foregroundOutcomeChainSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", foregroundOutcomeChainSource);
            Assert.DoesNotContain("ResolveCreatedCaseDisplaySession", foregroundOutcomeChainSource);
            Assert.DoesNotContain("_createdCaseDisplaySessions", foregroundOutcomeChainSource);
            Assert.DoesNotContain("IsCompleted", foregroundOutcomeChainSource);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", foregroundOutcomeChainSource);
            Assert.DoesNotContain("EvaluateCreatedCaseDisplayCompletionDecision", foregroundOutcomeChainSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", foregroundOutcomeChainSource);
        }

        [Fact]
        public void ForegroundTraceActionsSourcesAndDetails_PreserveContract()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string decisionTrace = ReadMethod(orchestrationSource, "private void LogForegroundGuaranteeDecision");
            string startedTrace = ReadMethod(orchestrationSource, "private void LogFinalForegroundGuaranteeStarted");
            string completedTrace = ReadMethod(orchestrationSource, "private void LogFinalForegroundGuaranteeCompleted");

            AssertContainsInOrder(
                decisionTrace,
                "action=foreground-recovery-decision",
                ", refreshSucceeded=",
                ", resolvedWindowPresent=",
                ", recoveryServicePresent=",
                ", foregroundRecoveryStarted=",
                ", foregroundRecoverySkipped=",
                ", foregroundSkipReason=",
                ", foregroundOutcomeStatus=",
                ", foregroundOutcomeDisplayCompletable=",
                "\"foreground-recovery-decision\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome\"",
                "\"reason=\" + (reason ?? string.Empty)",
                "\",foregroundRecoveryStarted=\" + foregroundRecoveryStarted.ToString()",
                "\",foregroundSkipReason=\" + (foregroundSkipReason ?? string.Empty)",
                "\",foregroundOutcomeStatus=\" + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString())");
            AssertContainsInOrder(
                startedTrace,
                "action=final-foreground-guarantee-start",
                "\"final-foreground-guarantee-started\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome\"",
                "\"reason=\" + (reason ?? string.Empty)");
            AssertContainsInOrder(
                completedTrace,
                "action=final-foreground-guarantee-end",
                ", recovered=",
                "\"final-foreground-guarantee-completed\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome\"",
                "\"reason=\" + (reason ?? string.Empty)",
                "\",recovered=\" + (executionResult != null && executionResult.Recovered).ToString()",
                "\",foregroundOutcomeStatus=\"",
                "? ForegroundGuaranteeOutcomeStatus.RequiredSucceeded.ToString()",
                ": ForegroundGuaranteeOutcomeStatus.RequiredDegraded.ToString()");
        }

        [Fact]
        public void PendingRetryAndActiveFallbackRefreshSuccessStopRetryWithoutCompletionOwnership()
        {
            string pendingSource = ReadAppSource("PendingPaneRefreshRetryService.cs");

            AssertContainsInOrder(
                pendingSource,
                "action=defer-retry-end",
                "refreshed=",
                "if (refreshed)",
                "_stopPendingPaneRefreshTimer();");
            AssertContainsInOrder(
                pendingSource,
                "action=defer-active-context-fallback-end",
                "refreshed=",
                "if (fallbackRefreshed)",
                "_stopPendingPaneRefreshTimer();");
            Assert.DoesNotContain("case-display-completed", pendingSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", pendingSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", pendingSource);
        }

        [Fact]
        public void ForegroundAndNormalizedOutcomesDoNotOwnCompletionEmit()
        {
            string attemptResultSource = ReadAppSource("TaskPaneRefreshAttemptResult.cs");
            string normalizedMapperSource = ReadAppSource("TaskPaneNormalizedOutcomeMapper.cs");
            string requiredDegraded = Slice(
                attemptResultSource,
                "internal static ForegroundGuaranteeOutcome RequiredDegraded",
                "internal static ForegroundGuaranteeOutcome RequiredFailed");

            Assert.Contains("ForegroundGuaranteeOutcomeStatus.RequiredDegraded", requiredDegraded);
            Assert.Contains("isTerminal: true", requiredDegraded);
            Assert.Contains("isDisplayCompletable: true", requiredDegraded);
            Assert.Contains("recoverySucceeded: false", requiredDegraded);
            Assert.DoesNotContain("case-display-completed", attemptResultSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", attemptResultSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", attemptResultSource);
            Assert.DoesNotContain("case-display-completed", normalizedMapperSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", normalizedMapperSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", normalizedMapperSource);
        }

        private static void AssertDoesNotOwnCompletion(string appFileName)
        {
            string source = ReadAppSource(appFileName);
            Assert.DoesNotContain("action=case-display-completed", source);
            Assert.DoesNotContain("\"case-display-completed\"", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", source);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", source);
        }

        private static void AssertContainsInOrder(string source, params string[] fragments)
        {
            int previousIndex = -1;
            foreach (string fragment in fragments)
            {
                int index = source.IndexOf(fragment, previousIndex + 1, StringComparison.Ordinal);
                Assert.True(
                    index > previousIndex,
                    "Expected to find '" + fragment + "' after index " + previousIndex.ToString() + ".");
                previousIndex = index;
            }
        }

        private static void AssertNormalizedOutcomeChainBefore(
            string source,
            string boundaryFragment,
            string orchestrationSource)
        {
            int boundaryIndex = source.IndexOf(boundaryFragment, StringComparison.Ordinal);
            Assert.True(boundaryIndex >= 0, "Expected boundary fragment was not found: " + boundaryFragment);

            int visibilityIndex = source.IndexOf("CompleteVisibilityRecoveryOutcome(", StringComparison.Ordinal);
            int refreshSourceIndex = source.IndexOf("CompleteRefreshSourceSelectionOutcome(", StringComparison.Ordinal);
            int rebuildFallbackIndex = source.IndexOf("CompleteRebuildFallbackOutcome(", StringComparison.Ordinal);
            if (visibilityIndex >= 0
                && refreshSourceIndex > visibilityIndex
                && rebuildFallbackIndex > refreshSourceIndex
                && rebuildFallbackIndex < boundaryIndex)
            {
                return;
            }

            int chainHelperIndex = source.IndexOf("CompleteNormalizedOutcomeChain(", StringComparison.Ordinal);
            Assert.True(
                chainHelperIndex >= 0 && chainHelperIndex < boundaryIndex,
                "Expected R10/R11/R12 direct chain or CompleteNormalizedOutcomeChain before " + boundaryFragment + ".");
            string helperSource = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteNormalizedOutcomeChain");
            AssertContainsInOrder(
                helperSource,
                "CompleteVisibilityRecoveryOutcome(",
                "CompleteRefreshSourceSelectionOutcome(",
                "CompleteRebuildFallbackOutcome(");
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", helperSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", helperSource);
        }

        private static string Slice(string source, string startFragment, string endFragment)
        {
            int start = source.IndexOf(startFragment, StringComparison.Ordinal);
            Assert.True(start >= 0, "Expected start fragment was not found: " + startFragment);
            int end = source.IndexOf(endFragment, start + startFragment.Length, StringComparison.Ordinal);
            Assert.True(end > start, "Expected end fragment was not found: " + endFragment);
            return source.Substring(start, end - start);
        }

        private static string ReadOptionalMethod(string source, string signatureFragment)
        {
            return source.IndexOf(signatureFragment, StringComparison.Ordinal) >= 0
                ? ReadMethod(source, signatureFragment)
                : string.Empty;
        }

        private static string ReadMethod(string source, string signatureFragment)
        {
            int start = source.IndexOf(signatureFragment, StringComparison.Ordinal);
            Assert.True(start >= 0, "Expected method signature was not found: " + signatureFragment);
            int openBrace = source.IndexOf('{', start);
            Assert.True(openBrace > start, "Expected method body was not found: " + signatureFragment);

            int depth = 0;
            for (int index = openBrace; index < source.Length; index++)
            {
                if (source[index] == '{')
                {
                    depth++;
                }
                else if (source[index] == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return source.Substring(start, index - start + 1);
                    }
                }
            }

            throw new InvalidOperationException("Method body was not closed: " + signatureFragment);
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
