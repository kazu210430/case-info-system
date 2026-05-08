using CaseInfoSystem.ExcelAddIn.App;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public sealed class WindowActivatePaneHandlingServiceTests
    {
        [Fact]
        public void Handle_WhenAllowed_DispatchesDisplayRefreshTriggerWithoutOwningRecovery()
        {
            WindowActivateTaskPaneTriggerFacts facts = CreateFacts(out Excel.Workbook workbook, out Excel.Window window);
            var predicateBridge = new FakeWindowActivatePanePredicateBridge();
            bool externalDetected = false;
            TaskPaneDisplayRequest dispatchedRequest = null;
            Excel.Workbook dispatchedWorkbook = null;
            Excel.Window dispatchedWindow = null;
            var service = new WindowActivatePaneHandlingService(
                predicateBridge,
                (_, __) => externalDetected = true,
                (_, __) => false,
                (request, targetWorkbook, targetWindow) =>
                {
                    dispatchedRequest = request;
                    dispatchedWorkbook = targetWorkbook;
                    dispatchedWindow = targetWindow;
                });

            WindowActivateDispatchOutcome outcome = service.Handle(facts);

            Assert.True(externalDetected);
            Assert.Same(workbook, dispatchedWorkbook);
            Assert.Same(window, dispatchedWindow);
            Assert.NotNull(dispatchedRequest);
            Assert.True(dispatchedRequest.IsWindowActivateTrigger);
            Assert.Equal("WindowActivate", dispatchedRequest.ToReasonString());
            Assert.Same(facts, dispatchedRequest.WindowActivateTriggerFacts);
            Assert.Equal(WindowActivateDispatchOutcomeStatus.Dispatched, outcome.Status);
            Assert.True(outcome.IsTerminal);
            Assert.Equal(WindowActivateActivationAttempt.NotAttempted, outcome.ActivationAttempt);
            Assert.False(outcome.IsDisplayCompletionOutcome);
            Assert.False(outcome.IsRecoveryOwner);
            Assert.False(outcome.IsForegroundGuaranteeOwner);
            Assert.False(outcome.IsHiddenExcelOwner);
            Assert.False(window.Activated);
        }

        [Fact]
        public void Handle_WhenProtected_IgnoresWithoutDispatching()
        {
            WindowActivateTaskPaneTriggerFacts facts = CreateFacts(out _, out _);
            var predicateBridge = new FakeWindowActivatePanePredicateBridge
            {
                ShouldIgnore = true
            };
            bool externalDetected = false;
            bool dispatched = false;
            var service = new WindowActivatePaneHandlingService(
                predicateBridge,
                (_, __) => externalDetected = true,
                (_, __) => false,
                (_, __, ___) => dispatched = true);

            WindowActivateDispatchOutcome outcome = service.Handle(facts);

            Assert.False(externalDetected);
            Assert.False(dispatched);
            Assert.Equal(WindowActivateDispatchOutcomeStatus.Ignored, outcome.Status);
            Assert.True(outcome.IsTerminal);
            Assert.Equal("caseProtection", outcome.OutcomeReason);
            Assert.Equal(WindowActivateActivationAttempt.NotAttempted, outcome.ActivationAttempt);
        }

        [Fact]
        public void Handle_WhenSuppressed_DefersWithoutDispatching()
        {
            WindowActivateTaskPaneTriggerFacts facts = CreateFacts(out _, out _);
            var predicateBridge = new FakeWindowActivatePanePredicateBridge();
            bool externalDetected = false;
            bool dispatched = false;
            var service = new WindowActivatePaneHandlingService(
                predicateBridge,
                (_, __) => externalDetected = true,
                (_, __) => true,
                (_, __, ___) => dispatched = true);

            WindowActivateDispatchOutcome outcome = service.Handle(facts);

            Assert.True(externalDetected);
            Assert.False(dispatched);
            Assert.Equal(WindowActivateDispatchOutcomeStatus.Deferred, outcome.Status);
            Assert.True(outcome.IsTerminal);
            Assert.Equal("casePaneRefreshSuppressed", outcome.OutcomeReason);
            Assert.Equal(WindowActivateActivationAttempt.NotAttempted, outcome.ActivationAttempt);
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
                "WindowActivatePaneHandlingServiceTests");
        }

        private sealed class FakeWindowActivatePanePredicateBridge : IWindowActivatePanePredicateBridge
        {
            internal bool ShouldIgnore { get; set; }

            public bool ShouldIgnoreDuringCaseProtection(Excel.Workbook workbook, Excel.Window window)
            {
                return ShouldIgnore;
            }
        }
    }
}
