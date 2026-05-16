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
            string payloadBuilderSource = ReadAppSource("TaskPaneRefreshEmitPayloadBuilder.cs");
            string completionMethod = ReadMethod(orchestrationSource, "private void TryCompleteCreatedCaseDisplaySession");
            string markMethod = ReadMethod(orchestrationSource, "private bool TryMarkCreatedCaseDisplaySessionCompletedForEmit");

            Assert.Contains("TryCompleteCreatedCaseDisplaySession", orchestrationSource);
            Assert.Contains("_emitPayloadBuilder.BuildCaseDisplayCompleted(", orchestrationSource);
            Assert.Contains("action=case-display-completed", payloadBuilderSource);
            Assert.Contains("\"case-display-completed\"", payloadBuilderSource);
            Assert.Contains("\"TaskPaneRefreshOrchestrationService.CompleteCreatedCaseDisplaySession\"", payloadBuilderSource);
            Assert.Contains("NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);", orchestrationSource);
            AssertContainsInOrder(
                completionMethod,
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))",
                "return;",
                "CaseDisplayCompletedPayload payload =");
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
            Assert.DoesNotContain("BuildCaseDisplayCompleted", markMethod);

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
            string completionDecisionSource = ReadAppSource("TaskPaneRefreshCompletionDecisionService.cs");

            Assert.Contains("IsCreatedCaseDisplayReason(reason)", orchestrationSource);
            Assert.Contains("_completionDecisionService.DecideCreatedCaseDisplayCompletion(", orchestrationSource);
            Assert.Contains("attemptResult == null", completionDecisionSource);
            Assert.Contains("!attemptResult.IsRefreshSucceeded", completionDecisionSource);
            Assert.Contains("!attemptResult.IsPaneVisible", completionDecisionSource);
            Assert.Contains("attemptResult.VisibilityRecoveryOutcome == null", completionDecisionSource);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsTerminal", completionDecisionSource);
            Assert.Contains("!attemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable", completionDecisionSource);
            Assert.Contains(
                "!IsForegroundDisplayCompletableTerminalInput(attemptResult.ForegroundGuaranteeOutcome)",
                completionDecisionSource);

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
            string payloadBuilderSource = ReadAppSource("TaskPaneRefreshEmitPayloadBuilder.cs");
            string payloadHelper = Slice(
                payloadBuilderSource,
                "internal CaseDisplayCompletedPayload BuildCaseDisplayCompleted",
                "return new CaseDisplayCompletedPayload");

            AssertContainsInOrder(
                completionMethod,
                "if (!completionDecision.CanComplete)",
                "CreatedCaseDisplaySession resolvedSession = session ?? ResolveCreatedCaseDisplaySession(reason, workbook);",
                "if (!TryMarkCreatedCaseDisplaySessionCompletedForEmit(resolvedSession))",
                "return;",
                "CaseDisplayCompletedPayload payload = _emitPayloadBuilder.BuildCaseDisplayCompleted(",
                "_logger?.Info(payload.KernelTraceMessage);",
                "payload.ObservationAction",
                "payload.ObservationSource",
                "payload.Details);",
                "NewCaseVisibilityObservation.Complete(resolvedSession.WorkbookFullName);");
            AssertContainsInOrder(
                payloadHelper,
                "string details =",
                "\"reason=\" + (input.Reason ?? string.Empty)",
                "\",sessionId=\" + input.SessionId",
                "\",completionSource=\" + (input.CompletionSource ?? string.Empty)",
                "\",completion=\" + input.AttemptResult.CompletionBasis",
                "\",paneVisible=\" + input.AttemptResult.IsPaneVisible.ToString()",
                "\",visibilityRecoveryStatus=\" + input.AttemptResult.VisibilityRecoveryOutcome.Status.ToString()",
                "\",visibilityRecoveryDisplayCompletable=\" + input.AttemptResult.VisibilityRecoveryOutcome.IsDisplayCompletable.ToString()",
                "\",visibilityRecoveryPaneVisible=\" + input.AttemptResult.VisibilityRecoveryOutcome.IsPaneVisible.ToString()",
                "\",visibilityRecoveryTargetKind=\" + input.AttemptResult.VisibilityRecoveryOutcome.TargetKind.ToString()",
                "\",visibilityPaneVisibleSource=\" + input.AttemptResult.VisibilityRecoveryOutcome.PaneVisibleSource.ToString()",
                "\",visibilityRecoveryReason=\" + input.AttemptResult.VisibilityRecoveryOutcome.Reason",
                "\",visibilityRecoveryDegradedReason=\" + input.AttemptResult.VisibilityRecoveryOutcome.DegradedReason",
                "\",refreshSourceStatus=\" + input.AttemptResult.RefreshSourceSelectionOutcome.Status.ToString()",
                "\",refreshSourceSelectedSource=\" + input.AttemptResult.RefreshSourceSelectionOutcome.SelectedSource.ToString()",
                "\",refreshSourceSelectionReason=\" + input.AttemptResult.RefreshSourceSelectionOutcome.SelectionReason",
                "\",refreshSourceFallbackReasons=\" + input.AttemptResult.RefreshSourceSelectionOutcome.FallbackReasons",
                "\",refreshSourceCacheFallback=\" + input.AttemptResult.RefreshSourceSelectionOutcome.IsCacheFallback.ToString()",
                "\",refreshSourceRebuildRequired=\" + input.AttemptResult.RefreshSourceSelectionOutcome.IsRebuildRequired.ToString()",
                "\",refreshSourceCanContinue=\" + input.AttemptResult.RefreshSourceSelectionOutcome.CanContinueRefresh.ToString()",
                "\",refreshSourceFailureReason=\" + input.AttemptResult.RefreshSourceSelectionOutcome.FailureReason",
                "\",refreshSourceDegradedReason=\" + input.AttemptResult.RefreshSourceSelectionOutcome.DegradedReason",
                "\",rebuildFallbackStatus=\" + input.AttemptResult.RebuildFallbackOutcome.Status.ToString()",
                "\",rebuildFallbackRequired=\" + input.AttemptResult.RebuildFallbackOutcome.IsRequired.ToString()",
                "\",rebuildFallbackCanContinue=\" + input.AttemptResult.RebuildFallbackOutcome.CanContinueRefresh.ToString()",
                "\",rebuildFallbackSnapshotSource=\" + input.AttemptResult.RebuildFallbackOutcome.SnapshotSource.ToString()",
                "\",rebuildFallbackReasons=\" + input.AttemptResult.RebuildFallbackOutcome.FallbackReasons",
                "\",rebuildFallbackFailureReason=\" + input.AttemptResult.RebuildFallbackOutcome.FailureReason",
                "\",rebuildFallbackDegradedReason=\" + input.AttemptResult.RebuildFallbackOutcome.DegradedReason",
                "\",refreshCompleted=\" + input.AttemptResult.IsRefreshCompleted.ToString()",
                "\",foregroundGuaranteeTerminal=\" + input.AttemptResult.IsForegroundGuaranteeTerminal.ToString()",
                "\",foregroundGuaranteeRequired=\" + input.AttemptResult.WasForegroundGuaranteeRequired.ToString()",
                "\",foregroundGuaranteeStatus=\" + input.AttemptResult.ForegroundGuaranteeOutcome.Status.ToString()",
                "\",foregroundGuaranteeDisplayCompletable=\" + input.AttemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable.ToString()",
                "\",foregroundGuaranteeExecutionAttempted=\" + input.AttemptResult.ForegroundGuaranteeOutcome.WasExecutionAttempted.ToString()",
                "\",foregroundGuaranteeTargetKind=\" + input.AttemptResult.ForegroundGuaranteeOutcome.TargetKind.ToString()",
                "\",foregroundRecoverySucceeded=\"",
                "\",foregroundOutcomeReason=\" + input.AttemptResult.ForegroundGuaranteeOutcome.Reason",
                "WindowActivateDownstreamObservation.FormatDisplayRequestTraceFields(input.DisplayRequest)",
                "details += \",attempt=\" + input.AttemptNumber.Value.ToString(CultureInfo.InvariantCulture);",
                "string kernelTraceMessage =");
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
                "CaseDisplayCompletedPayload payload = _emitPayloadBuilder.BuildCaseDisplayCompleted(",
                "_logger?.Info(payload.KernelTraceMessage);",
                "NewCaseVisibilityObservation.Log(",
                "payload.ObservationAction",
                "payload.ObservationSource",
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
                "CaseDisplayCompletedPayload payload = _emitPayloadBuilder.BuildCaseDisplayCompleted(");
            Assert.DoesNotContain("action=case-display-completed", beforePayloadBuild);
            Assert.DoesNotContain("\"case-display-completed\"", beforePayloadBuild);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Log", beforePayloadBuild);
            Assert.DoesNotContain("NewCaseVisibilityObservation.Complete", beforePayloadBuild);
            Assert.DoesNotContain("case-display-completed", markMethod);
            Assert.DoesNotContain("NewCaseVisibilityObservation", markMethod);
            Assert.DoesNotContain("BuildCaseDisplayCompleted", markMethod);

            Assert.Equal(1, CountOccurrences(completionMethod, "NewCaseVisibilityObservation.Complete("));
            Assert.Equal(1, CountOccurrences(completionMethod, "NewCaseVisibilityObservation.Log("));
            Assert.Equal(1, CountOccurrences(completionMethod, "payload.KernelTraceMessage"));
            Assert.Equal(1, CountOccurrences(completionMethod, "payload.ObservationAction"));
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
            Assert.DoesNotContain("BuildCaseDisplayCompleted", missingSessionGate);
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
                "string workbookFullName = SafeWorkbookFullName(workbook);",
                "lock (_createdCaseDisplaySessionSyncRoot)",
                "_createdCaseDisplaySessionStateReader.ResolveForCompletion(",
                "new CreatedCaseDisplaySessionResolutionInput(",
                "IsCreatedCaseDisplayReason(reason)",
                "workbookFullName",
                "SnapshotCreatedCaseDisplaySessions()",
                "if (resolvedSnapshot != null",
                "_createdCaseDisplaySessions.TryGetValue(resolvedSnapshot.WorkbookFullName, out CreatedCaseDisplaySession session)",
                "return session;",
                "return null;");
            Assert.DoesNotContain("case-display-completed", lookupMethod);
            Assert.DoesNotContain("NewCaseVisibilityObservation", lookupMethod);
            Assert.DoesNotContain("IsCompleted =", lookupMethod);
            Assert.DoesNotContain("_createdCaseDisplaySessions.Remove", lookupMethod);
            Assert.DoesNotContain("BuildCaseDisplayCompleted", lookupMethod);
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
                "_completionDecisionService.DecideCreatedCaseDisplayCompletion(",
                "if (!completionDecision.CanComplete)",
                "return;");
            Assert.DoesNotContain("ResolveCreatedCaseDisplaySession", hardGateBlock);
            Assert.DoesNotContain("_createdCaseDisplaySessions", hardGateBlock);
            Assert.DoesNotContain("IsCompleted", hardGateBlock);
            Assert.DoesNotContain("BuildCaseDisplayCompleted", hardGateBlock);
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
            string preconditionBoundary = ReadMethod(orchestrationSource, "private TaskPaneRefreshPreconditionDecisionResult EvaluateTaskPaneRefreshPreconditionBoundary");
            string failClosedReturn = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult ReturnFailClosedTaskPaneRefreshResult");
            string preconditionServiceSource = ReadAppSource("TaskPaneRefreshPreconditionDecisionService.cs");
            string preconditionSkipPath = Slice(
                coreMethod,
                "if (!preconditionDecision.CanRefresh)",
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(");

            AssertContainsInOrder(
                coreMethod,
                "TaskPaneRefreshAttemptStartObservation attemptObservation = StartTaskPaneRefreshAttemptObservation(",
                "TaskPaneRefreshPreconditionDecisionResult preconditionDecision = EvaluateTaskPaneRefreshPreconditionBoundary(",
                "if (!preconditionDecision.CanRefresh)",
                "return ReturnFailClosedTaskPaneRefreshResult(",
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(");
            AssertContainsInOrder(
                preconditionBoundary,
                "_preconditionDecisionService.Decide(",
                "if (!preconditionDecision.CanRefresh)",
                "_logger?.Info(",
                "return preconditionDecision;");
            AssertContainsInOrder(
                preconditionServiceSource,
                "TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(",
                "if (preconditionDecision.CanRefresh)",
                "return TaskPaneRefreshPreconditionDecisionResult.Continue(preconditionDecision);",
                "TaskPaneRefreshFailClosedOutcome.FromPreconditionDecision(preconditionDecision)");
            AssertNormalizedOutcomeChainBefore(
                failClosedReturn,
                "_windowActivateDownstreamObservation.LogOutcome(",
                orchestrationSource);
            AssertContainsInOrder(
                failClosedReturn,
                "TaskPaneRefreshAttemptResult skippedResult = CompleteNormalizedOutcomeChain(",
                "preconditionDecision.NormalizedOutcomeAttemptResult,",
                "preconditionDecision.NormalizedOutcomeActionName,",
                "_windowActivateDownstreamObservation.LogOutcome(",
                "preconditionDecision.SkipActionName);",
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
            string preconditionBoundary = ReadMethod(orchestrationSource, "private TaskPaneRefreshPreconditionDecisionResult EvaluateTaskPaneRefreshPreconditionBoundary");
            string routeDispatch = ReadMethod(orchestrationSource, "private RefreshDispatchExecutionResult DispatchTaskPaneRefreshRoute");
            string postDispatchConvergence = ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult ContinuePostDispatchRefreshConvergence");
            string dispatchShell = ReadMethod(orchestrationSource, "internal static RefreshDispatchExecutionResult Dispatch");

            AssertContainsInOrder(
                coreMethod,
                "TaskPaneRefreshAttemptStartObservation attemptObservation = StartTaskPaneRefreshAttemptObservation(",
                "TaskPaneRefreshPreconditionDecisionResult preconditionDecision = EvaluateTaskPaneRefreshPreconditionBoundary(",
                "if (!preconditionDecision.CanRefresh)",
                "return ReturnFailClosedTaskPaneRefreshResult(",
                "RefreshDispatchExecutionResult routeDispatchExecutionResult = DispatchTaskPaneRefreshRoute(",
                "return ContinuePostDispatchRefreshConvergence(");
            AssertContainsInOrder(
                preconditionBoundary,
                "_preconditionDecisionService.Decide(",
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
            string decisionServiceSource = ReadAppSource("TaskPaneRefreshObservationDecisionService.cs");
            string normalizedOutcomeChainSource =
                ReadMethod(orchestrationSource, "private TaskPaneRefreshAttemptResult CompleteNormalizedOutcomeChain")
                + ReadMethod(orchestrationSource, "private void LogVisibilityRecoveryOutcome")
                + ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionOutcome")
                + ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionTrace")
                + ReadMethod(orchestrationSource, "private void LogRebuildFallbackOutcome")
                + ReadMethod(decisionServiceSource, "internal TaskPaneRefreshObservationDecision CompleteNormalizedOutcomeChain");

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
            string decisionServiceSource = ReadAppSource("TaskPaneRefreshObservationDecisionService.cs");
            string visibilityTrace = ReadMethod(orchestrationSource, "private void LogVisibilityRecoveryOutcome");
            visibilityTrace = ReadMethod(decisionServiceSource, "internal static TaskPaneRefreshVisibilityObservationDecision Create") + visibilityTrace;
            string refreshSourceTrace = ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionOutcome")
                + ReadMethod(orchestrationSource, "private void LogRefreshSourceSelectionTrace");
            refreshSourceTrace = ReadMethod(decisionServiceSource, "internal static TaskPaneRefreshSourceObservationDecision Create") + refreshSourceTrace;
            string rebuildFallbackTrace = ReadMethod(orchestrationSource, "private void LogRebuildFallbackOutcome");
            rebuildFallbackTrace = ReadMethod(decisionServiceSource, "internal static TaskPaneRefreshRebuildFallbackObservationDecision Create") + rebuildFallbackTrace;

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
                "string statusAction = \"rebuild-fallback-\" + outcome.Status.ToString().ToLowerInvariant()",
                "\"rebuild-fallback-required\"",
                "\"TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome\"",
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
            string traceBuilderSource = ReadAppSource("TaskPaneForegroundGuaranteeTraceBuilder.cs");
            string decisionTrace = ReadMethod(traceBuilderSource, "internal TaskPaneForegroundGuaranteeTracePayload BuildDecisionTrace");
            string startedTrace = ReadMethod(traceBuilderSource, "internal TaskPaneForegroundGuaranteeTracePayload BuildStartedTrace");
            string completedTrace = ReadMethod(traceBuilderSource, "internal TaskPaneForegroundGuaranteeTracePayload BuildCompletedTrace");
            string orchestrationDecisionTrace = ReadMethod(orchestrationSource, "private void LogForegroundGuaranteeDecision");

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
                "ForegroundGuaranteeOutcomeSource");
            AssertContainsInOrder(
                startedTrace,
                "action=final-foreground-guarantee-start",
                "\"final-foreground-guarantee-started\"",
                "ForegroundGuaranteeOutcomeSource");
            AssertContainsInOrder(
                completedTrace,
                "action=final-foreground-guarantee-end",
                ", recovered=",
                "\"final-foreground-guarantee-completed\"",
                "ForegroundGuaranteeOutcomeSource");
            AssertContainsInOrder(
                orchestrationDecisionTrace,
                "_foregroundTraceBuilder.BuildDecisionTrace(",
                "_logger?.Info(trace.KernelTraceMessage);",
                "NewCaseVisibilityObservation.Log(",
                "trace.ObservationAction",
                "trace.ObservationSource",
                "trace.Details);");
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
                "\"reason=\" + (input.Reason ?? string.Empty)",
                "\",foregroundRecoveryStarted=\" + (input.Decision != null && input.Decision.ForegroundRecoveryStarted).ToString()",
                "\",foregroundSkipReason=\" + (input.Decision == null ? string.Empty : input.Decision.ForegroundSkipReason)",
                "\",foregroundOutcomeStatus=\" + (outcome == null ? ForegroundGuaranteeOutcomeStatus.Unknown.ToString() : outcome.Status.ToString())");
            Assert.DoesNotContain("foregroundOutcomeDisplayCompletable", decisionDetailsSource);
            Assert.DoesNotContain("recovered=", decisionDetailsSource);
            Assert.DoesNotContain("case-display-completed", decisionDetailsSource);

            Assert.Contains("\"reason=\" + (input.Reason ?? string.Empty)", startedDetailsSource);
            Assert.DoesNotContain("foregroundRecoveryStarted", startedDetailsSource);
            Assert.DoesNotContain("foregroundOutcomeStatus", startedDetailsSource);
            Assert.DoesNotContain("recovered=", startedDetailsSource);
            Assert.DoesNotContain("case-display-completed", startedDetailsSource);

            AssertContainsInOrder(
                completedDetailsSource,
                "\"reason=\" + (input.Reason ?? string.Empty)",
                "\",recovered=\" + recovered.ToString()",
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
                "bool recovered = input.ExecutionResult != null && input.ExecutionResult.Recovered;",
                "\",foregroundOutcomeStatus=\"",
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
            string workbookTargetRetry = ReadMethod(pendingSource, "private TaskPaneRefreshRetryContinuationDecision TryRefreshPendingWorkbookTarget");
            string activeContextFallback = ReadMethod(pendingSource, "private TaskPaneRefreshRetryContinuationDecision TryRefreshPendingActiveContextFallback");
            string retryContinuation = ReadMethod(pendingSource, "private void ResolvePendingRetryContinuation");

            AssertContainsInOrder(
                workbookTargetRetry,
                "action=defer-retry-end",
                "refreshed=",
                "return _retryContinuationDecisionService.DecideAfterRefresh(refreshed);");
            AssertContainsInOrder(
                activeContextFallback,
                "action=defer-active-context-fallback-end",
                "refreshed=",
                "return _retryContinuationDecisionService.DecideAfterRefresh(fallbackRefreshed);");
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
            string handoffFlow = ReadMethod(orchestrationSource, "private void RunTaskPaneRefreshHandoffFlow");
            string activeHandoffBranch = ReadMethod(orchestrationSource, "private void RunActiveTaskPaneRefreshHandoffBranch");
            string beginActiveHandoff = ReadMethod(orchestrationSource, "private ActiveTaskPaneRefreshHandoff BeginActiveTaskPaneRefreshHandoff");
            string immediateActiveRefresh = ReadMethod(orchestrationSource, "private bool TryRefreshActiveTaskPaneImmediately");
            string startActiveRetry = ReadMethod(orchestrationSource, "private void StartPendingRefreshRetryFromActiveHandoff");
            string activeRefreshHandoffFlow = scheduleActive
                + activeHandoffBranch
                + beginActiveHandoff
                + immediateActiveRefresh
                + startActiveRetry;

            AssertContainsInOrder(
                scheduleActive,
                "RunTaskPaneRefreshHandoffFlow(TaskPaneRefreshHandoffFlowInput.ForActiveRefresh(reason));");
            AssertContainsInOrder(
                handoffFlow,
                "if (flowInput.IsActiveRefresh)",
                "RunActiveTaskPaneRefreshHandoffBranch(flowInput);",
                "return;",
                "RunWorkbookFallbackTaskPaneRefreshHandoffBranch(flowInput);");
            AssertContainsInOrder(
                activeHandoffBranch,
                "ActiveTaskPaneRefreshHandoff activeHandoff = BeginActiveTaskPaneRefreshHandoff(flowInput.Reason);",
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
            string handoffFlow = ReadMethod(orchestrationSource, "private void RunTaskPaneRefreshHandoffFlow");
            string workbookFallbackBranch = ReadMethod(orchestrationSource, "private void RunWorkbookFallbackTaskPaneRefreshHandoffBranch");
            string beginFallbackHandoff = ReadMethod(orchestrationSource, "private PendingFallbackRefreshHandoff BeginPendingFallbackRefreshHandoff");
            string skipFallbackBoundary = ReadMethod(orchestrationSource, "private static bool ShouldSkipPendingFallbackForWorkbookOpenBoundary");
            string prepareRetryHandoff = ReadMethod(orchestrationSource, "private PendingRefreshRetryHandoff PreparePendingRefreshRetryHandoff");
            string immediateFallbackRefresh = ReadMethod(orchestrationSource, "private bool TryRefreshPendingFallbackImmediately");
            string startPendingRetry = ReadMethod(orchestrationSource, "private void StartPendingRefreshRetryFromFallback");
            string pendingFallbackHandoffFlow = scheduleFallback
                + handoffFlow
                + workbookFallbackBranch
                + beginFallbackHandoff
                + skipFallbackBoundary
                + prepareRetryHandoff
                + immediateFallbackRefresh
                + startPendingRetry;
            string workerShowWhenReady = ReadMethod(workerSource, "internal void ShowWhenReady");
            string schedulerSchedule = ReadMethod(schedulerSource, "internal void Schedule");
            string pendingBeginRetry = ReadMethod(pendingSource, "internal int BeginRetrySequence");
            string pendingTick = ReadMethod(pendingSource, "private void PendingPaneRefreshTimer_Tick");
            string pendingWorkbookTargetRetry = ReadMethod(pendingSource, "private TaskPaneRefreshRetryContinuationDecision TryRefreshPendingWorkbookTarget");
            string pendingActiveContextFallback = ReadMethod(pendingSource, "private TaskPaneRefreshRetryContinuationDecision TryRefreshPendingActiveContextFallback");
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
                "RunTaskPaneRefreshHandoffFlow(TaskPaneRefreshHandoffFlowInput.ForWorkbookFallback(workbook, reason));");
            AssertContainsInOrder(
                handoffFlow,
                "RunWorkbookFallbackTaskPaneRefreshHandoffBranch(flowInput);");
            AssertContainsInOrder(
                workbookFallbackBranch,
                "PendingFallbackRefreshHandoff fallbackHandoff = BeginPendingFallbackRefreshHandoff(flowInput.WorkbookFallbackWorkbook, flowInput.Reason);",
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
                "TaskPaneRefreshRetryContinuationDecision startDecision =",
                "_retryContinuationDecisionService.DecideBeforeTick(_retryState.HasAttemptsRemaining);",
                "if (startDecision.ShouldStopTimer)",
                "ResolvePendingRetryContinuation(startDecision);",
                "TaskPaneRefreshRetryContinuationDecision tickResult = TryRefreshPendingWorkbookTarget();",
                "if (!tickResult.Handled)",
                "tickResult = TryRefreshPendingActiveContextFallback();",
                "ResolvePendingRetryContinuation(tickResult);");
            AssertContainsInOrder(
                pendingWorkbookTargetRetry,
                "ResolvePendingPaneRefreshWorkbook()",
                "_retryContinuationDecisionService.DecideAfterWorkbookTargetResolution(targetWorkbook != null);",
                "if (!targetDecision.Handled)",
                "return targetDecision;",
                "_resolveWorkbookPaneWindow(targetWorkbook, _retryState.Reason, true);",
                "_tryRefreshTaskPane(_retryState.Reason, targetWorkbook, workbookWindow)",
                "return _retryContinuationDecisionService.DecideAfterRefresh(refreshed);");
            AssertContainsInOrder(
                pendingActiveContextFallback,
                "WorkbookContext context =",
                "_retryContinuationDecisionService.DecideActiveContextFallback(context);",
                "if (!contextDecision.ShouldAttemptActiveContextFallback)",
                "return contextDecision;",
                "bool fallbackRefreshed = _tryRefreshTaskPane(_retryState.Reason, null, null).IsRefreshSucceeded;",
                "return _retryContinuationDecisionService.DecideAfterRefresh(fallbackRefreshed);");
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
            string traceBuilderSource = ReadAppSource("TaskPaneForegroundGuaranteeTraceBuilder.cs");
            string method = ReadMethod(traceBuilderSource, "internal TaskPaneForegroundGuaranteeTracePayload BuildDecisionTrace");
            return Slice(method, "string details =", "string message =");
        }

        private static string ReadFinalForegroundStartedObservationDetailsSource(string source)
        {
            string traceBuilderSource = ReadAppSource("TaskPaneForegroundGuaranteeTraceBuilder.cs");
            string method = ReadMethod(traceBuilderSource, "internal TaskPaneForegroundGuaranteeTracePayload BuildStartedTrace");
            return Slice(method, "string details =", "string message =");
        }

        private static string ReadFinalForegroundCompletedObservationDetailsSource(string source)
        {
            string traceBuilderSource = ReadAppSource("TaskPaneForegroundGuaranteeTraceBuilder.cs");
            string method = ReadMethod(traceBuilderSource, "internal TaskPaneForegroundGuaranteeTracePayload BuildCompletedTrace");
            return Slice(method, "bool recovered =", "string message =");
        }

        private static string ReadForegroundExecutionClassificationSource(string source)
        {
            const string predicate = "executionResult.ExecutionAttempted && executionResult.Recovered";
            const string succeededFactory = "ForegroundGuaranteeOutcome.RequiredSucceeded";
            const string degradedFactory = "ForegroundGuaranteeOutcome.RequiredDegraded";

            int start = source.IndexOf(predicate, StringComparison.Ordinal);
            if (start < 0)
            {
                source = ReadAppSource("TaskPaneRefreshObservationDecisionService.cs");
                start = source.IndexOf(predicate, StringComparison.Ordinal);
            }
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

            string decisionServiceSource = ReadAppSource("TaskPaneRefreshObservationDecisionService.cs");
            helperSource =
                ReadOptionalMethod(decisionServiceSource, "internal bool IsForegroundDisplayCompletableTerminalInput")
                + ReadOptionalMethod(decisionServiceSource, "private static bool HasForegroundDisplayCompletableTerminalInput")
                + ReadOptionalMethod(decisionServiceSource, "private static bool HasDisplayCompletableForegroundOutcome")
                + ReadOptionalMethod(decisionServiceSource, "private static bool IsForegroundGuaranteeDisplayCompletableInput");
            if (!string.IsNullOrEmpty(helperSource))
            {
                return helperSource;
            }

            string completionDecisionSource = ReadAppSource("TaskPaneRefreshCompletionDecisionService.cs");
            helperSource =
                ReadOptionalMethod(completionDecisionSource, "internal static bool IsForegroundDisplayCompletableTerminalInput")
                + ReadOptionalMethod(completionDecisionSource, "private static bool HasForegroundDisplayCompletableTerminalInput")
                + ReadOptionalMethod(completionDecisionSource, "private static bool HasDisplayCompletableForegroundOutcome")
                + ReadOptionalMethod(completionDecisionSource, "private static bool IsForegroundGuaranteeDisplayCompletableInput");
            if (!string.IsNullOrEmpty(helperSource))
            {
                return helperSource;
            }

            return ReadMethod(completionDecisionSource, "internal CreatedCaseDisplayCompletionDecision DecideCreatedCaseDisplayCompletion");
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
            string decisionServiceSource = ReadAppSource("TaskPaneRefreshObservationDecisionService.cs");
            AssertContainsInOrder(
                helperSource,
                "_observationDecisionService.CompleteNormalizedOutcomeChain(",
                "LogVisibilityRecoveryOutcome(",
                "LogRefreshSourceSelectionOutcome(",
                "LogRebuildFallbackOutcome(");
            AssertContainsInOrder(
                decisionServiceSource,
                "TaskPaneNormalizedOutcomeMapper.BuildVisibilityRecoveryOutcome(",
                "TaskPaneNormalizedOutcomeMapper.BuildRefreshSourceSelectionOutcome(",
                "TaskPaneNormalizedOutcomeMapper.BuildRebuildFallbackOutcome(");
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
