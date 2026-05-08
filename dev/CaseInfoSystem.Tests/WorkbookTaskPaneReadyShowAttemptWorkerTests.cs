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
            Assert.False(shownOutcome.RefreshAttemptResult.IsRefreshCompleted);
            Assert.True(shownOutcome.RefreshAttemptResult.IsForegroundGuaranteeTerminal);
            Assert.Equal(
                ForegroundGuaranteeOutcomeStatus.SkippedAlreadyVisible,
                shownOutcome.RefreshAttemptResult.ForegroundGuaranteeOutcome.Status);
            Assert.True(shownOutcome.RefreshAttemptResult.ForegroundGuaranteeOutcome.IsDisplayCompletable);
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
            Assert.True(shownOutcome.RefreshAttemptResult.IsRefreshCompleted);
            Assert.Equal(
                ForegroundGuaranteeOutcomeStatus.NotRequired,
                shownOutcome.RefreshAttemptResult.ForegroundGuaranteeOutcome.Status);
        }

        private static WorkbookTaskPaneReadyShowAttemptWorker CreateWorker(
            int maxAttempts,
            Func<Excel.Workbook, Excel.Window, bool> hasVisibleCasePaneForWorkbookWindow,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            out Excel.Workbook workbook,
            out Excel.Window window)
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

            var logger = OrchestrationTestSupport.CreateLogger(new List<string>());
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
