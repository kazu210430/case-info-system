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
        public void HandleWorkbookBeforeClose_WhenSessionIsClean_LogsClosePreFactsBeforeSchedulingPostClose()
        {
            var loggerMessages = new List<string>();
            string scheduledKey = null;
            string scheduledFolder = null;
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(loggerMessages),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-immutable",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    SchedulePostCloseFollowUp = (key, folder) =>
                    {
                        scheduledKey = key;
                        scheduledFolder = folder;
                    }
                });

            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.False(handled);
            Assert.False(cancel);
            Assert.Equal("case-immutable", scheduledKey);
            Assert.Equal(@"C:\cases", scheduledFolder);
            Assert.Contains(
                loggerMessages,
                message => message.IndexOf("action=workbook-close-immutable-facts-captured", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("workbook=case-immutable", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("isBaseOrCaseWorkbook=True", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("isManagedClose=False", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("isSessionDirty=False", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("beforeCloseAction=SchedulePostCloseFollowUp", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Contains(
                loggerMessages,
                message => message.IndexOf("action=workbook-close-follow-up-facts-captured", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("workbook=case-immutable", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("folderPathCaptured=True", StringComparison.OrdinalIgnoreCase) >= 0
                    && message.IndexOf("beforeCloseAction=SchedulePostCloseFollowUp", StringComparison.OrdinalIgnoreCase) >= 0);
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

        [Fact]
        public void HandleWorkbookOpenedOrActivated_WhenInitializationFailsThenSucceeds_RetriesUntilSuccessThenNoOps()
        {
            var callLog = new List<string>();
            int syncAttempts = 0;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsCaseWorkbook = _ => true,
                    RegisterKnownCaseWorkbook = _ => callLog.Add("register"),
                    SyncNameRulesFromKernelToCase = _ =>
                    {
                        syncAttempts++;
                        callLog.Add("sync:" + syncAttempts.ToString());
                        if (syncAttempts == 1)
                        {
                            throw new InvalidOperationException("sync failed");
                        }
                    }
                });

            service.HandleWorkbookOpenedOrActivated(workbook);
            service.HandleWorkbookOpenedOrActivated(workbook);
            service.HandleWorkbookOpenedOrActivated(workbook);

            Assert.Equal(
                new[]
                {
                    "register",
                    "sync:1",
                    "register",
                    "sync:2"
                },
                callLog);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenPromptIsCanceled_DoesNotScheduleManagedCloseOrPostClose()
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
                    IsSuppressed = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    ShowClosePrompt = _ => DialogResult.Cancel,
                    ScheduleManagedSessionClose = (key, folder, saveChanges) => managedCloseCalls++,
                    SchedulePostCloseFollowUp = (key, folder) => postCloseCalls++
                });

            service.HandleSheetChanged(workbook);
            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.True(handled);
            Assert.True(cancel);
            Assert.Equal(0, managedCloseCalls);
            Assert.Equal(0, postCloseCalls);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenPromptThrows_DoesNotScheduleCloseAndRestoresCancel()
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
                    IsSuppressed = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    ShowClosePrompt = _ => throw new InvalidOperationException("prompt failed"),
                    ScheduleManagedSessionClose = (key, folder, saveChanges) => managedCloseCalls++,
                    SchedulePostCloseFollowUp = (key, folder) => postCloseCalls++
                });

            service.HandleSheetChanged(workbook);
            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.False(handled);
            Assert.False(cancel);
            Assert.Equal(0, managedCloseCalls);
            Assert.Equal(0, postCloseCalls);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenPromptThrows_ThenNextNormalCloseWorksCorrectly()
        {
            var callLog = new List<string>();
            int promptAttempts = 0;
            Excel.Workbook workbook = new Excel.Workbook();
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsSuppressed = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    ShowClosePrompt = _ =>
                    {
                        promptAttempts++;
                        callLog.Add("prompt:" + promptAttempts.ToString());
                        if (promptAttempts == 1)
                        {
                            throw new InvalidOperationException("prompt failed");
                        }

                        return DialogResult.Yes;
                    },
                    ScheduleManagedSessionClose = (key, folder, saveChanges) =>
                    {
                        callLog.Add("schedule:" + saveChanges.ToString());
                    },
                    SchedulePostCloseFollowUp = (key, folder) => callLog.Add("post-close")
                });

            service.HandleSheetChanged(workbook);

            bool firstCancel = false;
            bool firstHandled = service.HandleWorkbookBeforeClose(workbook, ref firstCancel);

            Assert.False(firstHandled);
            Assert.False(firstCancel);
            Assert.Equal(
                new[]
                {
                    "prompt:1"
                },
                callLog);

            bool secondCancel = false;
            bool secondHandled = service.HandleWorkbookBeforeClose(workbook, ref secondCancel);

            Assert.True(secondHandled);
            Assert.True(secondCancel);
            Assert.Equal(
                new[]
                {
                    "prompt:1",
                    "prompt:2",
                    "schedule:True"
                },
                callLog);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenCreatedCaseFolderOfferIsPending_PromptsBeforeSchedulingManagedClose()
        {
            var callLog = new List<string>();
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case-1.xlsx",
                Name = "case-1.xlsx",
                Path = @"C:\cases"
            };
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    IsSuppressed = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    DirectoryExistsSafe = _ => true,
                    ShowClosePrompt = _ => DialogResult.Yes,
                    ShowCreatedCaseFolderOfferPrompt = folder =>
                    {
                        callLog.Add("folder-prompt");
                        return DialogResult.Yes;
                    },
                    OpenCreatedCaseFolder = (folder, reason) => callLog.Add("folder-open"),
                    ScheduleManagedSessionClose = (key, folder, saveChanges) =>
                    {
                        callLog.Add("schedule:" + saveChanges.ToString());
                    },
                    SchedulePostCloseFollowUp = (key, folder) => callLog.Add("post-close")
                });

            service.MarkCreatedCaseFolderOfferPending(workbook);
            service.HandleSheetChanged(workbook);
            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.True(handled);
            Assert.True(cancel);
            Assert.Equal(new[] { "folder-prompt", "folder-open", "schedule:True" }, callLog);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenCreatedCaseFolderOfferIsPendingOnCleanClose_PromptsBeforeSchedulingPostClose()
        {
            var callLog = new List<string>();
            bool cancel = false;
            Excel.Workbook workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case-1.xlsx",
                Name = "case-1.xlsx",
                Path = @"C:\cases"
            };
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    DirectoryExistsSafe = _ => true,
                    ShowCreatedCaseFolderOfferPrompt = folder =>
                    {
                        callLog.Add("folder-prompt");
                        return DialogResult.No;
                    },
                    OpenCreatedCaseFolder = (folder, reason) => callLog.Add("folder-open"),
                    ScheduleManagedSessionClose = (key, folder, saveChanges) => callLog.Add("managed"),
                    SchedulePostCloseFollowUp = (key, folder) =>
                    {
                        callLog.Add("post-close");
                    }
                });

            service.MarkCreatedCaseFolderOfferPending(workbook);
            bool handled = service.HandleWorkbookBeforeClose(workbook, ref cancel);

            Assert.False(handled);
            Assert.False(cancel);
            Assert.Equal(new[] { "folder-prompt", "post-close" }, callLog);
        }

        [Fact]
        public void HandleWorkbookBeforeClose_WhenCreatedCaseFolderOfferPromptWasAlreadyShown_DoesNotPromptTwice()
        {
            var callLog = new List<string>();
            Excel.Workbook workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case-1.xlsx",
                Name = "case-1.xlsx",
                Path = @"C:\cases"
            };
            var service = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = _ => "case-1",
                    IsBaseOrCaseWorkbook = _ => true,
                    IsManagedClose = _ => false,
                    ResolveContainingFolder = _ => @"C:\cases",
                    DirectoryExistsSafe = _ => true,
                    ShowCreatedCaseFolderOfferPrompt = folder =>
                    {
                        callLog.Add("folder-prompt");
                        return DialogResult.No;
                    },
                    SchedulePostCloseFollowUp = (key, folder) => callLog.Add("post-close")
                });

            service.MarkCreatedCaseFolderOfferPending(workbook);

            bool firstCancel = false;
            bool secondCancel = false;
            service.HandleWorkbookBeforeClose(workbook, ref firstCancel);
            service.HandleWorkbookBeforeClose(workbook, ref secondCancel);

            Assert.Equal(
                new[]
                {
                    "folder-prompt",
                    "post-close",
                    "post-close"
                },
                callLog);
        }
    }
}
