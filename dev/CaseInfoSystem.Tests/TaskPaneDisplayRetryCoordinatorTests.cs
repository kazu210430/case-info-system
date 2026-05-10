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
            string completionMethod = ReadMethod(orchestrationSource, "private void TryCompleteCreatedCaseDisplaySession");
            string markMethod = ReadMethod(orchestrationSource, "private bool TryMarkCreatedCaseDisplaySessionCompletedForEmit");

            Assert.Contains("TryCompleteCreatedCaseDisplaySession", orchestrationSource);
            Assert.Contains("action=case-display-completed", orchestrationSource);
            Assert.Contains("\"case-display-completed\"", orchestrationSource);
            Assert.Contains("\"TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession\"", orchestrationSource);
            Assert.Contains("NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);", orchestrationSource);
            AssertContainsInOrder(
                completionMethod,
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))",
                "return;",
                "string details =");
            AssertContainsInOrder(
                markMethod,
                "bool shouldEmit = false;",
                "lock (_createdCaseDisplaySessionSyncRoot)",
                "if (!resolvedSession.IsCompleted)",
                "resolvedSession.IsCompleted = true;",
                "_createdCaseDisplaySessions.Remove(resolvedSession.WorkbookFullName);",
                "shouldEmit = true;",
                "return shouldEmit;");
            Assert.DoesNotContain("case-display-completed", markMethod);
            Assert.DoesNotContain("NewCaseVisibilityObservation", markMethod);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", markMethod);

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
            Assert.Contains(
                "!IsForegroundDisplayCompletableTerminalInput(attemptResult.ForegroundGuaranteeOutcome)",
                orchestrationSource);

            string foregroundInputSource = ReadOptionalForegroundDisplayCompletableDecisionSource(orchestrationSource);
            Assert.Contains("outcome != null", foregroundInputSource);
            Assert.Contains("outcome.IsTerminal", foregroundInputSource);
            Assert.Contains("outcome.IsDisplayCompletable", foregroundInputSource);
        }

        [Fact]
        public void CaseDisplayCompletedDetailsPayload_PreservesFieldSetAndOrder()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string completionMethod = ReadMethod(orchestrationSource, "private void TryCompleteCreatedCaseDisplaySession");
            string payloadHelper = Slice(
                orchestrationSource,
                "private static string BuildCaseDisplayCompletedDetailsPayload",
                "private static CreatedCaseDisplayCompletionDecision");

            AssertContainsInOrder(
                completionMethod,
                "if (!completionDecision.CanComplete)",
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);",
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))",
                "return;",
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
        public void CaseDisplayCompleted_OneTimeGuardBlocksDuplicateBeforeEmitAndComplete()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string completionMethod = ReadMethod(orchestrationSource, "private void TryCompleteCreatedCaseDisplaySession");
            string markMethod = ReadMethod(orchestrationSource, "private bool TryMarkCreatedCaseDisplaySessionCompletedForEmit");

            AssertContainsInOrder(
                completionMethod,
                "CreatedCaseDisplayCompletionDecision completionDecision =",
                "if (!completionDecision.CanComplete)",
                "return;",
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);",
                "if (resolvedSession == null)",
                "return;",
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))",
                "return;",
                "string details = BuildCaseDisplayCompletedDetailsPayload(",
                "_logger?.Info(",
                "action=case-display-completed",
                "NewCaseVisibilityObservation.Log(",
                "\"case-display-completed\"",
                "NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);");
            AssertContainsInOrder(
                markMethod,
                "bool shouldEmit = false;",
                "lock (_createdCaseDisplaySessionSyncRoot)",
                "if (!resolvedSession.IsCompleted)",
                "resolvedSession.IsCompleted = true;",
                "_createdCaseDisplaySessions.Remove(resolvedSession.WorkbookFullName);",
                "shouldEmit = true;",
                "return shouldEmit;");

            string beforePayloadBuild = Slice(
                completionMethod,
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))",
                "string details = BuildCaseDisplayCompletedDetailsPayload(");
            Assert.DoesNotContain("action=case-display-completed", beforePayloadBuild);
            Assert.DoesNotContain("\"case-display-completed\"", beforePayloadBuild);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Log", beforePayloadBuild);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", beforePayloadBuild);
            Assert.DoesNotContain("case-display-completed", markMethod);
            Assert.DoesNotContain("NewCaseVisibilityObservation", markMethod);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", markMethod);

            Assert.Equal(1, CountOccurrences(completionMethod, "NewCaseVisibilityObservation.Complete("));
            Assert.Equal(1, CountOccurrences(completionMethod, "NewCaseVisibilityObservation.Log("));
            Assert.Equal(1, CountOccurrences(completionMethod, "action=case-display-completed"));
            Assert.Equal(1, CountOccurrences(completionMethod, "\"case-display-completed\""));
            Assert.Equal(1, CountOccurrences(markMethod, "resolvedSession.IsCompleted = true;"));
            Assert.Equal(1, CountOccurrences(markMethod, "_createdCaseDisplaySessions.Remove(resolvedSession.WorkbookFullName);"));
            Assert.Equal(1, CountOccurrences(markMethod, "shouldEmit = true;"));
        }

        [Fact]
        public void CaseDisplayCompleted_MissingSessionReturnsBeforeOneTimeGuardEmitAndComplete()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string completionMethod = ReadMethod(orchestrationSource, "private void TryCompleteCreatedCaseDisplaySession");
            string missingSessionGate = Slice(
                completionMethod,
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);",
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))");

            AssertContainsInOrder(
                missingSessionGate,
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);",
                "if (resolvedSession == null)",
                "return;");
            Assert.DoesNotContain("resolvedSession.IsCompleted", missingSessionGate);
            Assert.DoesNotContain("_createdCaseDisplaySessions.Remove", missingSessionGate);
            Assert.DoesNotContain("TryMarkCreatedCaseDisplaySessionCompletedForEmit", missingSessionGate);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", missingSessionGate);
            Assert.DoesNotContain("action=case-display-completed", missingSessionGate);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Log", missingSessionGate);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", missingSessionGate);
        }

        [Fact]
        public void CaseDisplaySessionLookup_DoesNotOwnEmitCompletionOrSessionMutation()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string lookupMethod = ReadMethod(orchestrationSource, "private CreatedCaseDisplaySession ResolveCreatedCaseDisplaySession");

            AssertContainsInOrder(
                lookupMethod,
                "if (!IsCreatedCaseDisplayReason(reason))",
                "return null;",
                "string workbookFullName = SafeWorkbookFullName(workbook);",
                "lock (_createdCaseDisplaySessionSyncRoot)",
                "_createdCaseDisplaySessions.TryGetValue(workbookFullName, out CreatedCaseDisplaySession session)",
                "return session;",
                "if (_createdCaseDisplaySessions.Count == 1)",
                "return activeSession;",
                "return null;");
            Assert.DoesNotContain("case-display-completed", lookupMethod);
            Assert.DoesNotContain("NewCaseVisibilityObservation", lookupMethod);
            Assert.DoesNotContain("IsCompleted", lookupMethod);
            Assert.DoesNotContain("_createdCaseDisplaySessions.Remove", lookupMethod);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", lookupMethod);
        }

        [Fact]
        public void CaseDisplayCompletionHardGateBlocksBeforeSessionLookupMutationEmitAndComplete()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string completionMethod = ReadMethod(orchestrationSource, "private void TryCompleteCreatedCaseDisplaySession");
            string hardGateBlock = Slice(
                completionMethod,
                "CreatedCaseDisplayCompletionDecision completionDecision =",
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);");

            AssertContainsInOrder(
                hardGateBlock,
                "CreatedCaseDisplayCompletionDecision completionDecision =",
                "EvaluateCreatedCaseDisplayCompletionDecision(reason, attemptResult);",
                "if (!completionDecision.CanComplete)",
                "return;");
            Assert.DoesNotContain("ResolveCreatedCaseDisplaySession", hardGateBlock);
            Assert.DoesNotContain("_createdCaseDisplaySessions", hardGateBlock);
            Assert.DoesNotContain("IsCompleted", hardGateBlock);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", hardGateBlock);
            Assert.DoesNotContain("case-display-completed", hardGateBlock);
            Assert.DoesNotContain("NewCaseVisibilityObservation", hardGateBlock);
        }

        [Fact]
        public void NormalRefreshPath_KeepsNormalizedOutcomeChainBeforeForegroundWindowActivateAndCompletion()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string coreMethod = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult TryRefreshTaskPaneCore");
            string postDispatchConvergence = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult ContinuePostDispatchRefreshConvergence");

            AssertNormalizedOutcomeChainBefore(
                postDispatchConvergence,
                "CompleteForegroundGuaranteeOutcome(",
                orchestrationSource);
            AssertContainsInOrder(
                coreMethod,
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(",
                "return ContinuePostDispatchRefreshConvergence(");
            AssertContainsInOrder(
                postDispatchConvergence,
                "CompleteNormalizedOutcomeChain(",
                "routeDispatchExecutionResult.AttemptResult,",
                "CompleteForegroundGuaranteeOutcome(",
                "_windowActivateDownstreamObservation.LogOutcome(",
                "TryCompleteCreatedCaseDisplaySession(");
        }

        [Fact]
        public void PreconditionSkipPath_KeepsNormalizedOutcomeChainBeforeWindowActivateAndReturnWithoutForegroundOrCompletion()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string coreMethod = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult TryRefreshTaskPaneCore");
            string preconditionBoundary = ReadMethod(orchestrationSource, "private TaskPaneRefreshPreconditionDecision EvaluateTaskPaneRefreshPreconditionBoundary");
            string failClosedReturn = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult ReturnFailClosedTaskPaneRefreshResult");
            string preconditionSkipPath = Slice(
                coreMethod,
                "if (!preconditionDecision.CanRefresh)",
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(");

            AssertContainsInOrder(
                coreMethod,
                "TaskPaneRefreshAttemptStartObservation attemptObservation = StartTaskPaneRefreshAttemptObservation(",
                "TaskPaneRefreshPreconditionDecision preconditionDecision = EvaluateTaskPaneRefreshPreconditionBoundary(",
                "if (!preconditionDecision.CanRefresh)",
                "return ReturnFailClosedTaskPaneRefreshResult(",
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(");
            AssertContainsInOrder(
                preconditionBoundary,
                "TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(",
                "if (!preconditionDecision.CanRefresh)",
                "_logger?.Info(",
                "return preconditionDecision;");
            AssertNormalizedOutcomeChainBefore(
                failClosedReturn,
                "_windowActivateDownstreamObservation.LogOutcome(",
                orchestrationSource);
            AssertContainsInOrder(
                failClosedReturn,
                "TaskPaneRefreshFailClosedResultHandoff failClosedHandoff = BuildFailClosedTaskPaneRefreshResultHandoff(preconditionDecision);",
                "TaskPaneRefreshAttemptResult skippedResult = CompleteNormalizedOutcomeChain(",
                "failClosedHandoff.AttemptResult,",
                "failClosedHandoff.SkipActionName,",
                "_windowActivateDownstreamObservation.LogOutcome(",
                "failClosedHandoff.SkipActionName);",
                "return skippedResult;");
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", preconditionSkipPath);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", failClosedReturn);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", preconditionSkipPath);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", failClosedReturn);
            Assert.DoesNotContain("case-display-completed", preconditionSkipPath);
            Assert.DoesNotContain("case-display-completed", failClosedReturn);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", preconditionSkipPath);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", failClosedReturn);
            Assert.DoesNotContain("RefreshDispatchShell.Dispatch(", preconditionSkipPath);
            Assert.DoesNotContain("RefreshDispatchShell.Dispatch(", failClosedReturn);
            Assert.DoesNotContain("foreground-recovery-decision", preconditionSkipPath);
            Assert.DoesNotContain("foreground-recovery-decision", failClosedReturn);
            Assert.DoesNotContain("final-foreground-guarantee", preconditionSkipPath);
            Assert.DoesNotContain("final-foreground-guarantee", failClosedReturn);
            Assert.DoesNotContain("RefreshDispatchShell.Dispatch(", preconditionBoundary);
            Assert.DoesNotContain("CompleteNormalizedOutcomeChain(", preconditionBoundary);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", preconditionBoundary);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", preconditionBoundary);
            Assert.DoesNotContain("case-display-completed", preconditionBoundary);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", preconditionBoundary);
        }

        [Fact]
        public void RouteDispatchShell_StaysBetweenPreconditionAndNormalizedOutcomeWithoutCompletionOwnership()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string coreMethod = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult TryRefreshTaskPaneCore");
            string preconditionBoundary = ReadMethod(orchestrationSource, "private TaskPaneRefreshPreconditionDecision EvaluateTaskPaneRefreshPreconditionBoundary");
            string routeDispatch = ReadMethod(orchestrationSource, "private RefreshDispatchExecutionResult DispatchTaskPaneRefreshRoute");
            string postDispatchConvergence = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult ContinuePostDispatchRefreshConvergence");
            string dispatchShell = ReadMethod(orchestrationSource, "internal static RefreshDispatchExecutionResult Dispatch");

            AssertContainsInOrder(
                coreMethod,
                "TaskPaneRefreshAttemptStartObservation attemptObservation = StartTaskPaneRefreshAttemptObservation(",
                "TaskPaneRefreshPreconditionDecision preconditionDecision = EvaluateTaskPaneRefreshPreconditionBoundary(",
                "if (!preconditionDecision.CanRefresh)",
                "return ReturnFailClosedTaskPaneRefreshResult(",
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(",
                "return ContinuePostDispatchRefreshConvergence(");
            AssertContainsInOrder(
                preconditionBoundary,
                "TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(",
                "return preconditionDecision;");
            AssertContainsInOrder(
                routeDispatch,
                "RefreshDispatchExecutionResult dispatchExecutionResult = RefreshDispatchShell.Dispatch(",
                "_taskPaneRefreshCoordinator,",
                "reason,",
                "workbook,",
                "window,",
                "_getKernelHomeForm,",
                "_getTaskPaneRefreshSuppressionCount);",
                "return dispatchExecutionResult;");
            AssertContainsInOrder(
                postDispatchConvergence,
                "TaskPaneRefreshAttemptResult attemptResult = CompleteNormalizedOutcomeChain(",
                "attemptResult = CompleteForegroundGuaranteeOutcome(",
                "_windowActivateDownstreamObservation.LogOutcome(",
                "TryCompleteCreatedCaseDisplaySession(",
                "attemptResult,",
                "\"refresh\",");
            AssertContainsInOrder(
                dispatchShell,
                "taskPaneRefreshCoordinator.TryRefreshTaskPane(",
                "return RefreshDispatchExecutionResult.FromAttemptResult(attemptResult);");
            Assert.DoesNotContain("CompleteNormalizedOutcomeChain(", dispatchShell);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", dispatchShell);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", dispatchShell);
            Assert.DoesNotContain("case-display-completed", dispatchShell);
            Assert.DoesNotContain("NewCaseVisibilityObservation", dispatchShell);
            Assert.DoesNotContain("CompleteNormalizedOutcomeChain(", routeDispatch);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", routeDispatch);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", routeDispatch);
            Assert.DoesNotContain("case-display-completed", routeDispatch);
            Assert.DoesNotContain("NewCaseVisibilityObservation", routeDispatch);
            Assert.DoesNotContain("RefreshDispatchShell.Dispatch(", preconditionBoundary);
            Assert.DoesNotContain("CompleteNormalizedOutcomeChain(", preconditionBoundary);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", preconditionBoundary);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", preconditionBoundary);
            Assert.DoesNotContain("case-display-completed", preconditionBoundary);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", preconditionBoundary);
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
        public void ReadyShowCallbackRawFacts_ArePassedToNormalizedForegroundAndCompletionGateInputs()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string callbackHandler = Slice(
                orchestrationSource,
                "private void HandleWorkbookTaskPaneShown",
                "private void TryCompleteCreatedCaseDisplaySession");
            string callbackFactsBuilder = Slice(
                orchestrationSource,
                "private static ReadyShowCallbackFacts BuildReadyShowCallbackFacts",
                "private readonly struct ReadyShowCallbackFacts");

            AssertContainsInOrder(
                callbackHandler,
                "if (outcome == null)",
                "return;",
                "ReadyShowCallbackFacts callbackFacts = BuildReadyShowCallbackFacts(outcome);",
                "TaskPaneRefreshAttemptResult attemptResult = CompleteNormalizedOutcomeChain(",
                "callbackFacts.WorkbookWindow,",
                "callbackFacts.RefreshAttemptResult,",
                "\"ready-show-attempt\",",
                "callbackFacts.AttemptNumber,",
                "callbackFacts.WorkbookWindowEnsureFacts);",
                "attemptResult = CompleteForegroundGuaranteeOutcome(",
                "callbackFacts.WorkbookWindow,",
                "attemptResult,",
                "TryCompleteCreatedCaseDisplaySession(",
                "callbackFacts.WorkbookWindow,",
                "attemptResult,",
                "\"ready-show-attempt\",",
                "callbackFacts.AttemptNumber);");
            AssertContainsInOrder(
                callbackFactsBuilder,
                "return new ReadyShowCallbackFacts(",
                "outcome.WorkbookWindow,",
                "outcome.RefreshAttemptResult,",
                "outcome.AttemptNumber,",
                "outcome.WorkbookWindowEnsureFacts);");
            Assert.DoesNotContain("TaskPaneNormalizedOutcomeMapper.", callbackHandler);
            Assert.DoesNotContain("VisibilityRecoveryOutcome.", callbackHandler);
            Assert.DoesNotContain("ForegroundGuaranteeOutcome.", callbackHandler);
            Assert.DoesNotContain("EvaluateCreatedCaseDisplayCompletionDecision", callbackHandler);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", callbackHandler);
            Assert.DoesNotContain("action=case-display-completed", callbackHandler);
            Assert.DoesNotContain("\"case-display-completed\"", callbackHandler);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", callbackHandler);
            Assert.DoesNotContain("_createdCaseDisplaySessions", callbackHandler);
            Assert.DoesNotContain("IsCompleted", callbackHandler);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", callbackFactsBuilder);
            Assert.DoesNotContain("case-display-completed", callbackFactsBuilder);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", callbackFactsBuilder);
        }

        [Fact]
        public void ReadyShowCallbackOutcomeModel_RemainsRawFactsWithoutCompletionOwnership()
        {
            string displayAttemptResultSource = ReadAppSource("WorkbookTaskPaneDisplayAttemptResult.cs");
            string readyShowOutcomeSource = Slice(
                displayAttemptResultSource,
                "internal sealed class WorkbookTaskPaneReadyShowAttemptOutcome",
                "internal sealed class WorkbookWindowVisibilityEnsureFacts");

            AssertContainsInOrder(
                readyShowOutcomeSource,
                "internal WorkbookTaskPaneReadyShowAttemptOutcome(",
                "AttemptNumber = attemptNumber;",
                "WorkbookWindow = workbookWindow;",
                "RefreshAttemptResult = refreshAttemptResult ?? TaskPaneRefreshAttemptResult.Failed();",
                "VisibleCasePaneAlreadyShown = visibleCasePaneAlreadyShown;",
                "WorkbookWindowEnsureFacts = workbookWindowEnsureFacts;",
                "internal TaskPaneRefreshAttemptResult RefreshAttemptResult { get; }",
                "internal WorkbookWindowVisibilityEnsureFacts WorkbookWindowEnsureFacts { get; }",
                "return RefreshAttemptResult.IsRefreshSucceeded && RefreshAttemptResult.IsPaneVisible;",
                "internal WorkbookTaskPaneReadyShowAttemptOutcome WithWorkbookWindowEnsureFacts(",
                "return new WorkbookTaskPaneReadyShowAttemptOutcome(",
                "AttemptNumber,",
                "WorkbookWindow,",
                "RefreshAttemptResult,",
                "VisibleCasePaneAlreadyShown,",
                "workbookWindowEnsureFacts);");
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", readyShowOutcomeSource);
            Assert.DoesNotContain("EvaluateCreatedCaseDisplayCompletionDecision", readyShowOutcomeSource);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", readyShowOutcomeSource);
            Assert.DoesNotContain("case-display-completed", readyShowOutcomeSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation", readyShowOutcomeSource);
            Assert.DoesNotContain("_createdCaseDisplaySessions", readyShowOutcomeSource);
            Assert.DoesNotContain("IsCompleted", readyShowOutcomeSource);
        }

        [Fact]
        public void ReadyShowCallback_NullOutcomeOrWorkbookMissingReturnsBeforeCompletionInputs()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string callbackHandler = Slice(
                orchestrationSource,
                "private void HandleWorkbookTaskPaneShown",
                "private void TryCompleteCreatedCaseDisplaySession");
            string nullOutcomeGate = Slice(
                callbackHandler,
                "if (outcome == null)",
                "Stopwatch stopwatch = Stopwatch.StartNew();");
            string workerSource = ReadAppSource("WorkbookTaskPaneReadyShowAttemptWorker.cs");
            string workerShowWhenReady = ReadMethod(workerSource, "internal void ShowWhenReady");
            string nullWorkbookGate = Slice(
                workerShowWhenReady,
                "if (workbook == null)",
                "_logger?.Info(");

            AssertContainsInOrder(
                callbackHandler,
                "StopPendingPaneRefreshTimer();",
                "if (outcome == null)",
                "return;",
                "Stopwatch stopwatch = Stopwatch.StartNew();");
            Assert.DoesNotContain("CompleteNormalizedOutcomeChain(", nullOutcomeGate);
            Assert.DoesNotContain("CompleteForegroundGuaranteeOutcome(", nullOutcomeGate);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession(", nullOutcomeGate);
            Assert.DoesNotContain("case-display-completed", nullOutcomeGate);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", nullOutcomeGate);

            AssertContainsInOrder(
                workerShowWhenReady,
                "if (workbook == null)",
                "return;",
                "_logger?.Info(");
            Assert.DoesNotContain("_taskPaneDisplayRetryCoordinator.ShowWhenReady", nullWorkbookGate);
            Assert.DoesNotContain("onShown?.Invoke", nullWorkbookGate);
            Assert.DoesNotContain("TryShowWorkbookTaskPaneOnce", nullWorkbookGate);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", nullWorkbookGate);
            Assert.DoesNotContain("case-display-completed", nullWorkbookGate);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", nullWorkbookGate);
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
                "\"TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome\"");
            AssertContainsInOrder(
                startedTrace,
                "action=final-foreground-guarantee-start",
                "\"final-foreground-guarantee-started\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome\"");
            AssertContainsInOrder(
                completedTrace,
                "action=final-foreground-guarantee-end",
                ", recovered=",
                "\"final-foreground-guarantee-completed\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome\"");
        }

        [Fact]
        public void ForegroundTraceObservationDetails_PreserveFieldSetAndOrder()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string decisionDetailsSource = ReadForegroundDecisionObservationDetailsSource(orchestrationSource);
            string startedDetailsSource = ReadFinalForegroundStartedObservationDetailsSource(orchestrationSource);
            string completedDetailsSource = ReadFinalForegroundCompletedObservationDetailsSource(orchestrationSource);

            AssertContainsInOrder(
                decisionDetailsSource,
                "\"reason=\" + (reason ?? string.Empty)",
                "\",foregroundRecoveryStarted=\" + foregroundRecoveryStarted.ToString()",
                "\",foregroundSkipReason=\" + (foregroundSkipReason ?? string.Empty)",
                "\",foregroundOutcomeStatus=\" + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString())");
            Assert.DoesNotContain("foregroundOutcomeDisplayCompletable", decisionDetailsSource);
            Assert.DoesNotContain("recovered=", decisionDetailsSource);
            Assert.DoesNotContain("case-display-completed", decisionDetailsSource);

            Assert.Contains("\"reason=\" + (reason ?? string.Empty)", startedDetailsSource);
            Assert.DoesNotContain("foregroundRecoveryStarted", startedDetailsSource);
            Assert.DoesNotContain("foregroundOutcomeStatus", startedDetailsSource);
            Assert.DoesNotContain("recovered=", startedDetailsSource);
            Assert.DoesNotContain("case-display-completed", startedDetailsSource);

            AssertContainsInOrder(
                completedDetailsSource,
                "\"reason=\" + (reason ?? string.Empty)",
                "\",recovered=\" + (executionResult != null && executionResult.Recovered).ToString()",
                "\",foregroundOutcomeStatus=\"");
            Assert.DoesNotContain("foregroundRecoveryStarted", completedDetailsSource);
            Assert.DoesNotContain("foregroundSkipReason", completedDetailsSource);
            Assert.DoesNotContain("case-display-completed", completedDetailsSource);
        }

        [Fact]
        public void ForegroundTraceCompletedDetails_MapsRecoveredToSucceededOtherwiseDegraded()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string completedDetailsSource = ReadFinalForegroundCompletedObservationDetailsSource(orchestrationSource);

            AssertContainsInOrder(
                completedDetailsSource,
                "\",foregroundOutcomeStatus=\"",
                "executionResult != null && executionResult.Recovered",
                "? ForegroundGuaranteeOutcomeStatus.RequiredSucceeded.ToString()",
                ": ForegroundGuaranteeOutcomeStatus.RequiredDegraded.ToString()");
            Assert.DoesNotContain("ForegroundGuaranteeOutcomeStatus.RequiredFailed", completedDetailsSource);
            Assert.DoesNotContain("case-display-completed", completedDetailsSource);
        }

        [Fact]
        public void ForegroundTraceDetailsAssembly_DoesNotOwnExecutionWindowActivateOrCompletion()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string detailsSource =
                ReadForegroundDecisionObservationDetailsSource(orchestrationSource)
                + ReadFinalForegroundStartedObservationDetailsSource(orchestrationSource)
                + ReadFinalForegroundCompletedObservationDetailsSource(orchestrationSource);

            Assert.DoesNotContain("_taskPaneRefreshCoordinator", detailsSource);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", detailsSource);
            Assert.DoesNotContain("BeginPostForegroundProtection", detailsSource);
            Assert.DoesNotContain("WindowActivate", detailsSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", detailsSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", detailsSource);
            Assert.DoesNotContain("ResolveCreatedCaseDisplaySession", detailsSource);
            Assert.DoesNotContain("_createdCaseDisplaySessions", detailsSource);
            Assert.DoesNotContain("IsCompleted", detailsSource);
            Assert.DoesNotContain("case-display-completed", detailsSource);
        }

        [Fact]
        public void ForegroundExecutionResultClassification_RequiresAttemptedRecoveredForRequiredSucceeded()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string classificationSource = ReadForegroundExecutionClassificationSource(orchestrationSource);

            AssertContainsInOrder(
                classificationSource,
                "executionResult.ExecutionAttempted && executionResult.Recovered",
                "ForegroundGuaranteeOutcome.RequiredSucceeded(targetKind, \"foregroundRecoverySucceeded\")",
                "ForegroundGuaranteeOutcome.RequiredDegraded(targetKind, \"foregroundRecoveryReturnedFalse\")");
            Assert.DoesNotContain("ForegroundGuaranteeOutcome.RequiredFailed", classificationSource);
            Assert.DoesNotContain("ForegroundGuaranteeOutcome.NotRequired", classificationSource);
        }

        [Fact]
        public void ForegroundExecutionResultClassification_IsOutcomeMappingOnlyWithoutRuntimeOwners()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string classificationSource = ReadForegroundExecutionClassificationSource(orchestrationSource);

            Assert.DoesNotContain("_taskPaneRefreshCoordinator", classificationSource);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", classificationSource);
            Assert.DoesNotContain("BeginPostForegroundProtection", classificationSource);
            Assert.DoesNotContain("WindowActivate", classificationSource);
            Assert.DoesNotContain("LogForegroundGuarantee", classificationSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation", classificationSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", classificationSource);
            Assert.DoesNotContain("EvaluateCreatedCaseDisplayCompletionDecision", classificationSource);
            Assert.DoesNotContain("BuildCaseDisplayCompletedDetailsPayload", classificationSource);
            Assert.DoesNotContain("case-display-completed", classificationSource);
            Assert.DoesNotContain("_createdCaseDisplaySessions", classificationSource);
            Assert.DoesNotContain("IsCompleted", classificationSource);
        }

        [Fact]
        public void PendingRetryAndActiveFallbackRefreshSuccessStopRetryWithoutCompletionOwnership()
        {
            string pendingSource = ReadAppSource("PendingPaneRefreshRetryService.cs");
            string workbookTargetRetry = ReadMethod(pendingSource, "private PendingRetryTickResult TryRefreshPendingWorkbookTarget");
            string activeContextFallback = ReadMethod(pendingSource, "private PendingRetryTickResult TryRefreshPendingActiveContextFallback");
            string retryContinuation = ReadMethod(pendingSource, "private void ResolvePendingRetryContinuation");

            AssertContainsInOrder(
                workbookTargetRetry,
                "action=defer-retry-end",
                "refreshed=",
                "? PendingRetryTickResult.StopRetrySequence()",
                ": PendingRetryTickResult.ContinueRetrySequence();");
            AssertContainsInOrder(
                activeContextFallback,
                "action=defer-active-context-fallback-end",
                "refreshed=",
                "? PendingRetryTickResult.StopRetrySequence()",
                ": PendingRetryTickResult.ContinueRetrySequence();");
            AssertContainsInOrder(
                retryContinuation,
                "if (tickResult.ShouldStopTimer)",
                "_stopPendingPaneRefreshTimer();");
            Assert.DoesNotContain("case-display-completed", pendingSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", pendingSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", pendingSource);
        }

        [Fact]
        public void ActiveTaskPaneRefreshHandoffSchedulesPendingRetryWithoutCompletionOwnership()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string scheduleActive = ReadMethod(orchestrationSource, "internal void ScheduleActiveTaskPaneRefresh");
            string beginActiveHandoff = ReadMethod(orchestrationSource, "private ActiveTaskPaneRefreshHandoff BeginActiveTaskPaneRefreshHandoff");
            string immediateActiveRefresh = ReadMethod(orchestrationSource, "private bool TryRefreshActiveTaskPaneImmediately");
            string startActiveRetry = ReadMethod(orchestrationSource, "private void StartPendingRefreshRetryFromActiveHandoff");
            string activeRefreshHandoffFlow = scheduleActive
                + beginActiveHandoff
                + immediateActiveRefresh
                + startActiveRetry;

            AssertContainsInOrder(
                scheduleActive,
                "ActiveTaskPaneRefreshHandoff activeHandoff = BeginActiveTaskPaneRefreshHandoff(reason);",
                "TryRefreshActiveTaskPaneImmediately(activeHandoff)",
                "return;",
                "StartPendingRefreshRetryFromActiveHandoff(activeHandoff);");
            AssertContainsInOrder(
                activeRefreshHandoffFlow,
                "_pendingPaneRefreshRetryService.TrackActiveTarget();",
                "IsTaskPaneRefreshSucceeded(activeHandoff.Reason, null, null)",
                "action=defer-immediate-success",
                ", target=active",
                "StopPendingPaneRefreshTimer();",
                "BeginRetrySequence(activeHandoff.Reason);",
                "action=defer-scheduled",
                ", target=active",
                ", attempts=");
            Assert.DoesNotContain("TrackWorkbookTarget", activeRefreshHandoffFlow);
            Assert.DoesNotContain("ResolveWorkbookPaneWindow", activeRefreshHandoffFlow);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", activeRefreshHandoffFlow);
            Assert.DoesNotContain("case-display-completed", activeRefreshHandoffFlow);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", activeRefreshHandoffFlow);
        }

        [Fact]
        public void ReadyShowPendingFallbackAndRetrySequencing_PreservesAttemptsDelayAndActivationMatrix()
        {
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string workerSource = ReadAppSource("WorkbookTaskPaneReadyShowAttemptWorker.cs");
            string schedulerSource = ReadAppSource("TaskPaneReadyShowRetryScheduler.cs");
            string pendingSource = ReadAppSource("PendingPaneRefreshRetryService.cs");
            string showWhenReady = ReadMethod(orchestrationSource, "internal void ShowWorkbookTaskPaneWhenReady");
            string scheduleFallback = ReadMethod(orchestrationSource, "internal void ScheduleWorkbookTaskPaneRefresh");
            string beginFallbackHandoff = ReadMethod(orchestrationSource, "private PendingFallbackRefreshHandoff BeginPendingFallbackRefreshHandoff");
            string skipFallbackBoundary = ReadMethod(orchestrationSource, "private static bool ShouldSkipPendingFallbackForWorkbookOpenBoundary");
            string prepareRetryHandoff = ReadMethod(orchestrationSource, "private PendingRefreshRetryHandoff PreparePendingRefreshRetryHandoff");
            string immediateFallbackRefresh = ReadMethod(orchestrationSource, "private bool TryRefreshPendingFallbackImmediately");
            string startPendingRetry = ReadMethod(orchestrationSource, "private void StartPendingRefreshRetryFromFallback");
            string pendingFallbackHandoffFlow = scheduleFallback
                + beginFallbackHandoff
                + skipFallbackBoundary
                + prepareRetryHandoff
                + immediateFallbackRefresh
                + startPendingRetry;
            string workerShowWhenReady = ReadMethod(workerSource, "internal void ShowWhenReady");
            string schedulerSchedule = ReadMethod(schedulerSource, "internal void Schedule");
            string pendingBeginRetry = ReadMethod(pendingSource, "internal int BeginRetrySequence");
            string pendingTick = ReadMethod(pendingSource, "private void PendingPaneRefreshTimer_Tick");
            string pendingWorkbookTargetRetry = ReadMethod(pendingSource, "private PendingRetryTickResult TryRefreshPendingWorkbookTarget");
            string pendingActiveContextFallback = ReadMethod(pendingSource, "private PendingRetryTickResult TryRefreshPendingActiveContextFallback");
            string pendingRetryContinuation = ReadMethod(pendingSource, "private void ResolvePendingRetryContinuation");

            Assert.Contains("internal const int ReadyShowMaxAttempts = 2;", workerSource);
            Assert.Contains("internal const int ReadyShowRetryDelayMs = 80;", workerSource);
            Assert.Contains("internal const int PendingPaneRefreshIntervalMs = 400;", orchestrationSource);
            Assert.Contains("internal const int PendingPaneRefreshMaxAttempts = 3;", orchestrationSource);
            AssertContainsInOrder(
                showWhenReady,
                "BeginCreatedCaseDisplaySession(workbook, reason)",
                "_workbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(",
                "_readyShowRetryScheduler.Schedule,",
                "outcome => HandleWorkbookTaskPaneShown(createdCaseDisplaySession, workbook, reason, outcome)",
                "ScheduleWorkbookTaskPaneRefresh");
            AssertContainsInOrder(
                workerShowWhenReady,
                "_taskPaneDisplayRetryCoordinator.ShowWhenReady(",
                "TryShowWorkbookTaskPaneOnce(targetWorkbook, targetReason, attemptNumber)",
                "scheduleRetry,",
                "() => onShown?.Invoke(shownOutcome)",
                "scheduleFallback");
            AssertContainsInOrder(
                schedulerSchedule,
                "_retryDelayMs.ToString(CultureInfo.InvariantCulture)",
                "if (retryAction == null)",
                "return;",
                "_retryTimerLifecycle.ScheduleWaitReadyRetryTimer(",
                "_retryDelayMs,",
                "retryAction();");
            AssertContainsInOrder(
                scheduleFallback,
                "PendingFallbackRefreshHandoff fallbackHandoff = BeginPendingFallbackRefreshHandoff(workbook, reason);",
                "ShouldSkipPendingFallbackForWorkbookOpenBoundary(fallbackHandoff)",
                "LogPendingFallbackWorkbookOpenSkip(fallbackHandoff);",
                "return;",
                "PendingRefreshRetryHandoff retryHandoff = PreparePendingRefreshRetryHandoff(fallbackHandoff);",
                "TryRefreshPendingFallbackImmediately(retryHandoff)",
                "return;",
                "StartPendingRefreshRetryFromFallback(retryHandoff);");
            AssertContainsInOrder(
                pendingFallbackHandoffFlow,
                "action=wait-ready-fallback-handoff",
                "ready-show-fallback-handoff",
                "TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(",
                "window: null",
                "_pendingPaneRefreshRetryService.TrackWorkbookTarget(",
                "ResolveWorkbookPaneWindow(workbook, reason, activateWorkbook: false);",
                "TryRefreshTaskPane(retryHandoff.Reason, retryHandoff.Workbook, retryHandoff.WorkbookWindow)",
                "StopPendingPaneRefreshTimer();",
                "BeginRetrySequence(retryHandoff.Reason);");
            AssertContainsInOrder(
                pendingBeginRetry,
                "_retryState.BeginRetrySequence(reason, _pendingPaneRefreshMaxAttempts);",
                "_retryTimerLifecycle.StartPendingPaneRefreshTimer(",
                "_pendingPaneRefreshIntervalMs,",
                "PendingPaneRefreshTimer_Tick);");
            AssertContainsInOrder(
                pendingTick,
                "if (!_retryState.HasAttemptsRemaining)",
                "ResolvePendingRetryContinuation(PendingRetryTickResult.StopRetrySequence());",
                "PendingRetryTickResult tickResult = TryRefreshPendingWorkbookTarget();",
                "if (!tickResult.Handled)",
                "tickResult = TryRefreshPendingActiveContextFallback();",
                "ResolvePendingRetryContinuation(tickResult);");
            AssertContainsInOrder(
                pendingWorkbookTargetRetry,
                "ResolvePendingPaneRefreshWorkbook()",
                "return PendingRetryTickResult.ContinueToActiveContextFallback();",
                "_resolveWorkbookPaneWindow(targetWorkbook, _retryState.Reason, true);",
                "_tryRefreshTaskPane(_retryState.Reason, targetWorkbook, workbookWindow)",
                "? PendingRetryTickResult.StopRetrySequence()",
                ": PendingRetryTickResult.ContinueRetrySequence();");
            AssertContainsInOrder(
                pendingActiveContextFallback,
                "WorkbookContext context =",
                "return PendingRetryTickResult.StopRetrySequence();",
                "bool fallbackRefreshed = _tryRefreshTaskPane(_retryState.Reason, null, null).IsRefreshSucceeded;",
                "? PendingRetryTickResult.StopRetrySequence()",
                ": PendingRetryTickResult.ContinueRetrySequence();");
            AssertContainsInOrder(
                pendingRetryContinuation,
                "if (tickResult.ShouldStopTimer)",
                "_stopPendingPaneRefreshTimer();");
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", pendingFallbackHandoffFlow);
            Assert.DoesNotContain("case-display-completed", pendingFallbackHandoffFlow);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", pendingFallbackHandoffFlow);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", pendingSource);
            Assert.DoesNotContain("case-display-completed", pendingSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", pendingSource);
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

        [Fact]
        public void ForegroundDisplayCompletableInputContract_IsMappingOnlyWithoutRuntimeOwners()
        {
            string attemptResultSource = ReadAppSource("TaskPaneRefreshAttemptResult.cs");
            string orchestrationSource = ReadAppSource("TaskPaneRefreshOrchestrationService.cs");
            string mappingSource =
                ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome Unknown")
                + ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome NotRequired")
                + ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome SkippedAlreadyVisible")
                + ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome SkippedNoKnownTarget")
                + ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome RequiredSucceeded")
                + ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome RequiredDegraded")
                + ReadMethod(attemptResultSource, "internal static ForegroundGuaranteeOutcome RequiredFailed")
                + ReadOptionalForegroundDisplayCompletableDecisionSource(orchestrationSource);

            Assert.Contains("isTerminal: true", mappingSource);
            Assert.Contains("isDisplayCompletable: true", mappingSource);
            Assert.Contains("isDisplayCompletable: false", mappingSource);
            Assert.Contains("ForegroundGuaranteeOutcomeStatus.RequiredDegraded", mappingSource);
            Assert.Contains("ForegroundGuaranteeOutcomeStatus.RequiredFailed", mappingSource);
            Assert.DoesNotContain("_taskPaneRefreshCoordinator", mappingSource);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", mappingSource);
            Assert.DoesNotContain("BeginPostForegroundProtection", mappingSource);
            Assert.DoesNotContain("WindowActivate", mappingSource);
            Assert.DoesNotContain("_logger", mappingSource);
            Assert.DoesNotContain("NewCaseVisibilityObservation", mappingSource);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", mappingSource);
            Assert.DoesNotContain("ResolveCreatedCaseDisplaySession", mappingSource);
            Assert.DoesNotContain("_createdCaseDisplaySessions", mappingSource);
            Assert.DoesNotContain("IsCompleted", mappingSource);
            Assert.DoesNotContain("case-display-completed", mappingSource);
        }

        private static void AssertDoesNotOwnCompletion(string appFileName)
        {
            string source = ReadAppSource(appFileName);
            Assert.DoesNotContain("action=case-display-completed", source);
            Assert.DoesNotContain("\"case-display-completed\"", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", source);
            Assert.DoesNotContain("TryCompleteCreatedCaseDisplaySession", source);
        }

        private static string ReadForegroundDecisionObservationDetailsSource(string source)
        {
            string helperSource = ReadOptionalMethod(source, "private static string BuildForegroundRecoveryDecisionDetails")
                + ReadOptionalMethod(source, "private static string BuildForegroundRecoveryDecisionObservationDetails");
            if (!string.IsNullOrEmpty(helperSource))
            {
                return helperSource;
            }

            string methodSource = ReadMethod(source, "private void LogForegroundGuaranteeDecision");
            return Slice(
                methodSource,
                "\"reason=\" + (reason ?? string.Empty)",
                ");");
        }

        private static string ReadFinalForegroundStartedObservationDetailsSource(string source)
        {
            string helperSource = ReadOptionalMethod(source, "private static string BuildFinalForegroundGuaranteeStartedDetails")
                + ReadOptionalMethod(source, "private static string BuildFinalForegroundGuaranteeStartedObservationDetails");
            if (!string.IsNullOrEmpty(helperSource))
            {
                return helperSource;
            }

            string methodSource = ReadMethod(source, "private void LogFinalForegroundGuaranteeStarted");
            return Slice(
                methodSource,
                "\"reason=\" + (reason ?? string.Empty)",
                ");");
        }

        private static string ReadFinalForegroundCompletedObservationDetailsSource(string source)
        {
            string helperSource = ReadOptionalMethod(source, "private static string BuildFinalForegroundGuaranteeCompletedDetails")
                + ReadOptionalMethod(source, "private static string BuildFinalForegroundGuaranteeCompletedObservationDetails");
            if (!string.IsNullOrEmpty(helperSource))
            {
                return helperSource;
            }

            string methodSource = ReadMethod(source, "private void LogFinalForegroundGuaranteeCompleted");
            return Slice(
                methodSource,
                "\"reason=\" + (reason ?? string.Empty)",
                ");");
        }

        private static string ReadForegroundExecutionClassificationSource(string source)
        {
            const string predicate = "executionResult.ExecutionAttempted && executionResult.Recovered";
            const string succeededFactory = "ForegroundGuaranteeOutcome.RequiredSucceeded";
            const string degradedFactory = "ForegroundGuaranteeOutcome.RequiredDegraded";

            int start = source.IndexOf(predicate, StringComparison.Ordinal);
            Assert.True(start >= 0, "Expected foreground classification predicate was not found.");
            int succeeded = source.IndexOf(succeededFactory, start, StringComparison.Ordinal);
            Assert.True(succeeded > start, "Expected RequiredSucceeded classification branch was not found.");
            int degraded = source.IndexOf(degradedFactory, succeeded, StringComparison.Ordinal);
            Assert.True(degraded > succeeded, "Expected RequiredDegraded classification branch was not found.");
            int end = source.IndexOf(';', degraded);
            Assert.True(end > degraded, "Expected foreground classification branch to end with a semicolon.");

            return source.Substring(start, end - start + 1);
        }

        private static string ReadOptionalForegroundDisplayCompletableDecisionSource(string source)
        {
            string helperSource =
                ReadOptionalMethod(source, "private static bool IsForegroundDisplayCompletableTerminalInput")
                + ReadOptionalMethod(source, "private static bool HasForegroundDisplayCompletableTerminalInput")
                + ReadOptionalMethod(source, "private static bool HasDisplayCompletableForegroundOutcome")
                + ReadOptionalMethod(source, "private static bool IsForegroundGuaranteeDisplayCompletableInput");
            if (!string.IsNullOrEmpty(helperSource))
            {
                return helperSource;
            }

            return ReadMethod(source, "private static CreatedCaseDisplayCompletionDecision EvaluateCreatedCaseDisplayCompletionDecision");
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

        private static int CountOccurrences(string source, string fragment)
        {
            int count = 0;
            int index = 0;
            while (true)
            {
                index = source.IndexOf(fragment, index, StringComparison.Ordinal);
                if (index < 0)
                {
                    return count;
                }

                count++;
                index += fragment.Length;
            }
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
