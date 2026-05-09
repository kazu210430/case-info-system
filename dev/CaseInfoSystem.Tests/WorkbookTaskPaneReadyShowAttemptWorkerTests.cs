using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    [Collection("ExcelApplicationCreatedApplications")]
    public sealed class WorkbookTaskPaneReadyShowAttemptWorkerTests : IDisposable
    {
        public WorkbookTaskPaneReadyShowAttemptWorkerTests()
        {
            Excel.Application.ResetCreatedApplications();
        }

        public void Dispose()
        {
            Excel.Application.ResetCreatedApplications();
        }

        [Fact]
        public void ShowWhenReady_WhenVisibleCasePaneAlreadyShown_CompletesWithoutRefresh()
        {
            int refreshCallCount = 0;
            var worker = CreateWorker(
                maxAttempts: 2,
                hasVisibleCasePaneForWorkbookWindow: (_, __) => true,
                tryRefreshTaskPane: (_, __, ___) =>
                {
                    refreshCallCount++;
                    return TaskPaneRefreshAttemptResult.Failed();
                },
                resolveWorkbookPaneWindow: (workbook, _, __) => workbook.Windows[1],
                out Excel.Workbook workbook,
                out _);

            bool shown = false;
            WorkbookTaskPaneReadyShowAttemptOutcome shownOutcome = null;
            bool fallbackScheduled = false;
            var scheduledAttempts = new List<int>();

            worker.ShowWhenReady(
                workbook,
                "ready",
                (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                outcome =>
                {
                    shown = true;
                    shownOutcome = outcome;
                },
                (_, __) => fallbackScheduled = true);

            Assert.True(shown);
            Assert.False(fallbackScheduled);
            Assert.Empty(scheduledAttempts);
            Assert.Equal(0, refreshCallCount);
            Assert.NotNull(shownOutcome);
            Assert.True(shownOutcome.VisibleCasePaneAlreadyShown);
            Assert.True(shownOutcome.IsShown);
            Assert.NotNull(shownOutcome.WorkbookWindowEnsureFacts);
            Assert.Equal(
                WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible,
                shownOutcome.WorkbookWindowEnsureFacts.Outcome);
            Assert.False(shownOutcome.RefreshAttemptResult.IsRefreshCompleted);
            Assert.True(shownOutcome.RefreshAttemptResult.IsForegroundGuaranteeTerminal);
            Assert.Equal(
                ForegroundGuaranteeOutcomeStatus.SkippedAlreadyVisible,
                shownOutcome.RefreshAttemptResult.ForegroundGuaranteeOutcome.Status);
            Assert.True(shownOutcome.RefreshAttemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable);
        }

        [Fact]
        public void ShowWhenReady_WhenVisibleCasePaneAlreadyShown_InvokesCallbackWithoutCompletionTrace()
        {
            var logs = new List<string>();
            int refreshCallCount = 0;
            var worker = CreateWorker(
                maxAttempts: 2,
                hasVisibleCasePaneForWorkbookWindow: (_, __) => true,
                tryRefreshTaskPane: (_, __, ___) =>
                {
                    refreshCallCount++;
                    return TaskPaneRefreshAttemptResult.Failed();
                },
                resolveWorkbookPaneWindow: (workbook, _, __) => workbook.Windows[1],
                out Excel.Workbook workbook,
                out _,
                logs: logs);

            WorkbookTaskPaneReadyShowAttemptOutcome shownOutcome = null;
            bool fallbackScheduled = false;

            worker.ShowWhenReady(
                workbook,
                "ready",
                (_, __, ___, ____) => throw new InvalidOperationException("ready-show retry should not run"),
                outcome => shownOutcome = outcome,
                (_, __) => fallbackScheduled = true);

            Assert.NotNull(shownOutcome);
            Assert.True(shownOutcome.IsShown);
            Assert.True(shownOutcome.VisibleCasePaneAlreadyShown);
            Assert.Equal("visibleCasePaneAlreadyShown", shownOutcome.RefreshAttemptResult.CompletionBasis);
            Assert.Equal(0, refreshCallCount);
            Assert.False(fallbackScheduled);
            Assert.False(logs.Exists(entry => entry.Contains("case-display-completed")));
        }

        [Fact]
        public void ShowWhenReady_WhenRetryRuns_DoesNotEnsureVisibilityAfterFirstAttempt()
        {
            bool? firstAttemptWindowVisible = null;
            bool? secondAttemptWindowVisible = null;
            int refreshCallCount = 0;
            var worker = CreateWorker(
                maxAttempts: 2,
                hasVisibleCasePaneForWorkbookWindow: (_, __) => false,
                tryRefreshTaskPane: (_, __, window) =>
                {
                    refreshCallCount++;
                    if (refreshCallCount == 1)
                    {
                        firstAttemptWindowVisible = window.Visible;
                        window.Visible = false;
                        return TaskPaneRefreshAttemptResult.Failed();
                    }

                    secondAttemptWindowVisible = window.Visible;
                    return TaskPaneRefreshAttemptResult.Succeeded();
                },
                resolveWorkbookPaneWindow: (workbook, _, __) => workbook.Windows[1],
                out Excel.Workbook workbook,
                out Excel.Window window);
            window.Visible = false;

            bool shown = false;
            WorkbookTaskPaneReadyShowAttemptOutcome shownOutcome = null;
            bool fallbackScheduled = false;
            var scheduledAttempts = new List<int>();

            worker.ShowWhenReady(
                workbook,
                "retry",
                (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                outcome =>
                {
                    shown = true;
                    shownOutcome = outcome;
                },
                (_, __) => fallbackScheduled = true);

            Assert.True(shown);
            Assert.False(fallbackScheduled);
            Assert.Equal(new[] { 2 }, scheduledAttempts);
            Assert.Equal(2, refreshCallCount);
            Assert.True(firstAttemptWindowVisible.GetValueOrDefault());
            Assert.False(secondAttemptWindowVisible.GetValueOrDefault());
            Assert.NotNull(shownOutcome);
            Assert.False(shownOutcome.VisibleCasePaneAlreadyShown);
            Assert.True(shownOutcome.IsShown);
            Assert.NotNull(shownOutcome.WorkbookWindowEnsureFacts);
            Assert.Equal(
                WorkbookWindowVisibilityEnsureOutcome.MadeVisible,
                shownOutcome.WorkbookWindowEnsureFacts.Outcome);
            Assert.True(shownOutcome.RefreshAttemptResult.IsRefreshCompleted);
            Assert.Equal(
                ForegroundGuaranteeOutcomeStatus.NotRequired,
                shownOutcome.RefreshAttemptResult.ForegroundGuaranteeOutcome.Status);
        }

        [Fact]
        public void ShowWhenReady_WhenReadyShowAttemptsFail_HandsOffToFallbackWithoutAttempt3()
        {
            int refreshCallCount = 0;
            var attemptedRefreshes = new List<int>();
            var worker = CreateWorker(
                maxAttempts: WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts,
                hasVisibleCasePaneForWorkbookWindow: (_, __) => false,
                tryRefreshTaskPane: (_, __, ___) =>
                {
                    refreshCallCount++;
                    attemptedRefreshes.Add(refreshCallCount);
                    return TaskPaneRefreshAttemptResult.Failed();
                },
                resolveWorkbookPaneWindow: (workbook, _, __) => workbook.Windows[1],
                out Excel.Workbook workbook,
                out _);

            bool shown = false;
            bool fallbackScheduled = false;
            string fallbackReason = null;
            var scheduledAttempts = new List<int>();

            worker.ShowWhenReady(
                workbook,
                "ready-fallback",
                (_, __, attemptNumber, continueAction) =>
                {
                    scheduledAttempts.Add(attemptNumber);
                    continueAction();
                },
                _ => shown = true,
                (_, reason) =>
                {
                    fallbackScheduled = true;
                    fallbackReason = reason;
                });

            Assert.False(shown);
            Assert.True(fallbackScheduled);
            Assert.Equal("ready-fallback", fallbackReason);
            Assert.Equal(new[] { 2 }, scheduledAttempts);
            Assert.Equal(new[] { 1, 2 }, attemptedRefreshes);
            Assert.Equal(2, refreshCallCount);
            Assert.DoesNotContain(3, scheduledAttempts);
            Assert.DoesNotContain(3, attemptedRefreshes);
        }

        private static WorkbookTaskPaneReadyShowAttemptWorker CreateWorker(
            int maxAttempts,
            Func<Excel.Workbook, Excel.Window, bool> hasVisibleCasePaneForWorkbookWindow,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            out Excel.Workbook workbook,
            out Excel.Window window,
            List<string> logs = null)
        {
            var application = new Excel.Application();
            workbook = new Excel.Workbook
            {
                Application = application,
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
            };
            application.Workbooks.Add(workbook);
            window = workbook.Windows[1];
            window.Hwnd = 101;
            window.Visible = true;
            application.ActiveWorkbook = workbook;
            application.ActiveWindow = window;

            var logger = OrchestrationTestSupport.CreateLogger(logs ?? new List<string>());
            var excelInteropService = new ExcelInteropService(application, logger, new PathCompatibilityService());
            var workbookWindowVisibilityService = new WorkbookWindowVisibilityService(excelInteropService, logger);
            return new WorkbookTaskPaneReadyShowAttemptWorker(
                excelInteropService,
                logger,
                new TaskPaneDisplayRetryCoordinator(maxAttempts),
                new WorkbookTaskPaneDisplayAttemptCoordinator(),
                workbookWindowVisibilityService,
                hasVisibleCasePaneForWorkbookWindow,
                tryRefreshTaskPane,
                resolveWorkbookPaneWindow);
        }
    }
}
