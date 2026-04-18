using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class CaseWorkbookLifecycleServiceThinOrchestrationTests
    {
        [Fact]
        public void HandleWorkbookOpenedOrActivated_WhenSameKeyIsSeenTwice_RunsInitializationOnlyOnce()
        {
            int registerCalls = 0;
            int syncCalls = 0;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsCaseWorkbook = _ => true,
                    RegisterKnownCaseWorkbook = _ => registerCalls++,
                    SyncNameRulesFromKernelToCase = _ => syncCalls++
                });

            service.HandleWorkbookOpenedOrActivated(workbook);
            service.HandleWorkbookOpenedOrActivated(workbook);

            Assert.Equal(1, registerCalls);
            Assert.Equal(1, syncCalls);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenManagedCloseIsActive_DoesNotPromptOrSchedule()
        {
            int promptCalls = 0;
            int managedCloseCalls = 0;
            int postCloseCalls = 0;
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => true,
                    ShowClosePrompt = _ =>
                    {
                        promptCalls++;
                        return DialogResult.Yes;
                    },
                    ScheduleManagedSessionClose = (key, folder, saveChanges) => managedCloseCalls++,
                    SchedulePostCloseFollowUp = (key, folder) => postCloseCalls++
                });

            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.False(handled);
            Assert.False(cancel);
            Assert.Equal(0, promptCalls);
            Assert.Equal(0, managedCloseCalls);
            Assert.Equal(0, postCloseCalls);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenSessionIsDirty_PromptsBeforeSchedulingManagedClose()
        {
            var callLog = new List<string>();
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    IsSuppressed = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    ShowClosePrompt = _ =>
                    {
                        callLog.Add("prompt");
                        return DialogResult.Yes;
                    },
                    ScheduleManagedSessionClose = (key, folder, saveChanges) =>
                    {
                        callLog.Add("schedule:" + saveChanges.ToString());
                    },
                    SchedulePostCloseFollowUp = (key, folder) => callLog.Add("post-close")
                });

            service.HandleSheetChanged(workbook);
            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.True(handled);
            Assert.True(cancel);
            Assert.Equal(
                new[]
                {
                    "prompt",
                    "schedule:True"
                },
                callLog);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenSessionIsClean_SchedulesOnlyPostCloseFollowUp()
        {
            int promptCalls = 0;
            int managedCloseCalls = 0;
            int postCloseCalls = 0;
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    ShowClosePrompt = _ =>
                    {
                        promptCalls++;
                        return DialogResult.Yes;
                    },
                    ScheduleManagedSessionClose = (key, folder, saveChanges) => managedCloseCalls++,
                    SchedulePostCloseFollowUp = (key, folder) => postCloseCalls++
                });

            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.False(handled);
            Assert.False(cancel);
            Assert.Equal(0, promptCalls);
            Assert.Equal(0, managedCloseCalls);
            Assert.Equal(1, postCloseCalls);
        }

        [Fact]
        public void HandleSheetChanged_WhenSuppressed_DoesNotDirtyWorkbook()
        {
            int managedCloseCalls = 0;
            int postCloseCalls = 0;
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    IsSuppressed = _ => true,
                    ResolveContainingFolder = _ => @"C:\cases",
                    ScheduleManagedSessionClose = (key, folder, saveChanges) => managedCloseCalls++,
                    SchedulePostCloseFollowUp = (key, folder) => postCloseCalls++
                });

            service.HandleSheetChanged(workbook);
            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.False(handled);
            Assert.False(cancel);
            Assert.Equal(0, managedCloseCalls);
            Assert.Equal(1, postCloseCalls);
        }

        [Fact]
        public void HandleWorkbookOpenedOrActivated_WhenInitializationThrows_RetriesOnNextActivation()
        {
            int registerCalls = 0;
            int syncCalls = 0;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsCaseWorkbook = _ => true,
                    RegisterKnownCaseWorkbook = _ => registerCalls++,
                    SyncNameRulesFromKernelToCase = _ =>
                    {
                        syncCalls++;
                        throw new InvalidOperationException("sync failed");
                    }
                });

            service.HandleWorkbookOpenedOrActivated(workbook);
            service.HandleWorkbookOpenedOrActivated(workbook);

            Assert.Equal(2, registerCalls);
            Assert.Equal(2, syncCalls);
        }
    }
}
