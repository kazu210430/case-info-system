using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookServiceThinOrchestrationTests
    {
        [Fact]
        public void ResolveKernelWorkbook_WhenContextIsNull_ReturnsPrimaryWithoutFallbackLookup()
        {
            int fallbackResolveCalls = 0;
            int fallbackFindCalls = 0;
            Excel.Workbook primary = new Excel.Workbook();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => primary,
                    ResolveKernelWorkbookPath = root =>
                    {
                        fallbackResolveCalls++;
                        return root + "\\kernel.xlsm";
                    },
                    FindOpenWorkbook = path =>
                    {
                        fallbackFindCalls++;
                        return null;
                    }
                });

            Excel.Workbook resolved = service.ResolveKernelWorkbook((WorkbookContext)null);

            Assert.Same(primary, resolved);
            Assert.Equal(0, fallbackResolveCalls);
            Assert.Equal(0, fallbackFindCalls);
        }

        [Fact]
        public void ResolveKernelWorkbook_WhenContextContainsKernelWorkbook_ReturnsContextWorkbookWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            int fallbackResolveCalls = 0;
            int fallbackFindCalls = 0;
            Excel.Workbook primary = new Excel.Workbook();
            Excel.Workbook contextKernelWorkbook = new Excel.Workbook { Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm") };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return primary;
                    },
                    ResolveKernelWorkbookPath = root =>
                    {
                        fallbackResolveCalls++;
                        return root + "\\kernel.xlsm";
                    },
                    FindOpenWorkbook = path =>
                    {
                        fallbackFindCalls++;
                        return null;
                    }
                });

            Excel.Workbook resolved = service.ResolveKernelWorkbook(
                new WorkbookContext(contextKernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", "kernel.xlsm", "shMasterList"));

            Assert.Same(contextKernelWorkbook, resolved);
            Assert.Equal(0, openKernelCalls);
            Assert.Equal(0, fallbackResolveCalls);
            Assert.Equal(0, fallbackFindCalls);
        }

        [Fact]
        public void ResolveKernelWorkbook_WhenMatchingRootWorkbookExists_ReturnsFallbackWorkbook()
        {
            Excel.Workbook fallback = new Excel.Workbook();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => new Excel.Workbook(),
                    ResolveKernelWorkbookPath = root => root + "\\kernel.xlsm",
                    FindOpenWorkbook = path => path == @"C:\root\kernel.xlsm" ? fallback : null
                });

            Excel.Workbook resolved = service.ResolveKernelWorkbook(new WorkbookContext(null, null, WorkbookRole.Case, @"C:\root", "case.xlsx", "shHOME"));

            Assert.Same(fallback, resolved);
        }

        [Fact]
        public void ResolveKernelWorkbook_WhenMatchingRootWorkbookDoesNotExist_ReturnsNull()
        {
            int openKernelCalls = 0;
            int fallbackFindCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    },
                    ResolveKernelWorkbookPath = root => string.Empty,
                    FindOpenWorkbook = path =>
                    {
                        fallbackFindCalls++;
                        return null;
                    }
                });

            Excel.Workbook resolved = service.ResolveKernelWorkbook(new WorkbookContext(null, null, WorkbookRole.Case, @"C:\root", "case.xlsx", "shHOME"));

            Assert.Null(resolved);
            Assert.Equal(0, openKernelCalls);
            Assert.Equal(0, fallbackFindCalls);
        }

        [Fact]
        public void CloseHomeSession_WhenNoOtherWorkbookExists_QuitsWithoutRestore()
        {
            int quitCalls = 0;
            var releaseCalls = new List<bool>();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => null,
                    HasOtherVisibleWorkbook = _ => false,
                    HasOtherWorkbook = _ => false,
                    ReleaseHomeDisplay = showExcel => releaseCalls.Add(showExcel),
                    QuitApplication = () => quitCalls++
                });

            service.CloseHomeSession();

            Assert.Single(releaseCalls);
            Assert.False(releaseCalls[0]);
            Assert.Equal(1, quitCalls);
        }

        [Fact]
        public void PrepareForHomeDisplay_WhenCalledTwice_AppliesVisibilityOnlyOnce()
        {
            int applyVisibilityCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => applyVisibilityCalls++
                });

            service.PrepareForHomeDisplay();
            service.PrepareForHomeDisplay();

            Assert.Equal(1, applyVisibilityCalls);
        }

        [Fact]
        public void PrepareForHomeDisplay_WhenHomeSessionCompletes_CanPrepareAgain()
        {
            var callLog = new List<string>();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => callLog.Add("apply"),
                    GetOpenKernelWorkbook = () => null,
                    HasOtherVisibleWorkbook = _ => false,
                    HasOtherWorkbook = _ => false,
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => callLog.Add("quit")
                });

            service.PrepareForHomeDisplay();
            service.CloseHomeSession();
            service.PrepareForHomeDisplay();

            Assert.Equal(
                new[]
                {
                    "apply",
                    "release:False",
                    "quit"
                },
                callLog);
        }

        [Fact]
        public void ResolveKernelWorkbook_WhenAvailabilityChanges_TransitionsFromPrimaryToFallbackToUnavailable()
        {
            var callLog = new List<string>();
            int phase = 0;
            Excel.Workbook primary = new Excel.Workbook();
            Excel.Workbook fallback = new Excel.Workbook();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        callLog.Add("get:" + phase.ToString());
                        return phase == 0 ? primary : null;
                    },
                    ResolveKernelWorkbookPath = root =>
                    {
                        callLog.Add("resolve:" + phase.ToString());
                        return phase == 2 ? string.Empty : root + "\\kernel.xlsm";
                    },
                    FindOpenWorkbook = path =>
                    {
                        callLog.Add("find:" + phase.ToString());
                        return phase == 1 ? fallback : null;
                    }
                });
            var context = new WorkbookContext(null, null, WorkbookRole.Case, @"C:\root", "case.xlsx", "shHOME");

            Excel.Workbook first = service.ResolveKernelWorkbook((WorkbookContext)null);
            phase = 1;
            Excel.Workbook second = service.ResolveKernelWorkbook(context);
            phase = 2;
            Excel.Workbook third = service.ResolveKernelWorkbook(context);

            Assert.Same(primary, first);
            Assert.Same(fallback, second);
            Assert.Null(third);
            Assert.Equal(
                new[]
                {
                    "get:0",
                    "resolve:1",
                    "find:1",
                    "resolve:2"
                },
                callLog);
        }

        [Fact]
        public void CloseHomeSessionSavingKernel_WhenCaseCreationFlowIsActive_ConcealsBeforeSaveAndDismissesWithoutRestoreOrQuit()
        {
            int quitCalls = 0;
            var callLog = new List<string>();
            var stateLogs = new List<string>();
            KernelCaseInteractionState interactionState = OrchestrationTestSupport.CreateKernelCaseInteractionState(stateLogs);
            using (interactionState.BeginKernelCaseCreationFlow("test"))
            {
                Excel.Workbook kernelWorkbook = new Excel.Workbook { FullName = @"C:\root\kernel.xlsm", Name = "kernel.xlsm" };
                var service = new KernelWorkbookService(
                    interactionState,
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    new KernelWorkbookService.KernelWorkbookServiceTestHooks
                    {
                        GetOpenKernelWorkbook = () => kernelWorkbook,
                        HasOtherVisibleWorkbook = _ => true,
                        HasOtherWorkbook = _ => true,
                        ApplyHomeDisplayVisibility = () => { },
                        ConcealKernelWorkbookWindowsForCaseCreationClose = workbook => callLog.Add("conceal"),
                        SaveAndCloseKernelWorkbook = workbook => callLog.Add("save-close"),
                        ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                        DismissPreparedHomeDisplayState = reason => callLog.Add("dismiss"),
                        QuitApplication = () => quitCalls++
                    });

                service.PrepareForHomeDisplay();
                service.CloseHomeSessionSavingKernel();
            }

            Assert.Equal(0, quitCalls);
            Assert.Equal(
                new[]
                {
                    "conceal",
                    "save-close",
                    "dismiss"
                },
                callLog);
        }

        [Fact]
        public void CloseHomeSession_WhenOtherWorkbookExists_RestoresWithoutQuit()
        {
            int quitCalls = 0;
            int dismissCalls = 0;
            var releaseCalls = new List<bool>();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => null,
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    ReleaseHomeDisplay = showExcel => releaseCalls.Add(showExcel),
                    DismissPreparedHomeDisplayState = reason => dismissCalls++,
                    QuitApplication = () => quitCalls++
                });

            service.CloseHomeSession();

            Assert.Single(releaseCalls);
            Assert.True(releaseCalls[0]);
            Assert.Equal(0, quitCalls);
            Assert.Equal(0, dismissCalls);
        }

        [Fact]
        public void CloseHomeSession_WhenManagedCloseIsRejected_DoesNotReleaseOrQuit()
        {
            int quitCalls = 0;
            var callLog = new List<string>();
            Excel.Workbook kernelWorkbook = new Excel.Workbook { FullName = @"C:\root\kernel.xlsm", Name = "kernel.xlsm" };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => kernelWorkbook,
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    RequestManagedCloseFromHomeExit = workbook =>
                    {
                        callLog.Add("request-close");
                        return false;
                    },
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => quitCalls++
                });

            service.SetLifecycleService(new KernelWorkbookLifecycleService());
            service.CloseHomeSession();

            Assert.Equal(0, quitCalls);
            Assert.Equal(new[] { "request-close" }, callLog);
        }

        [Fact]
        public void CloseHomeSession_WhenManagedCloseIsRetriedAfterRejection_ReusesPreparedState()
        {
            var callLog = new List<string>();
            int requestCalls = 0;
            Excel.Workbook kernelWorkbook = new Excel.Workbook { FullName = @"C:\root\kernel.xlsm", Name = "kernel.xlsm" };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => callLog.Add("apply"),
                    GetOpenKernelWorkbook = () => kernelWorkbook,
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    RequestManagedCloseFromHomeExit = workbook =>
                    {
                        requestCalls++;
                        callLog.Add("request:" + requestCalls.ToString());
                        return requestCalls > 1;
                    },
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => callLog.Add("quit")
                });

            service.SetLifecycleService(new KernelWorkbookLifecycleService());

            service.PrepareForHomeDisplay();
            service.CloseHomeSession();
            service.CloseHomeSession();

            Assert.Equal(
                new[]
                {
                    "apply",
                    "request:1",
                    "request:2",
                    "release:True"
                },
                callLog);
        }

        private static KernelWorkbookService CreateService(KernelWorkbookService.KernelWorkbookServiceTestHooks hooks)
        {
            return new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                hooks);
        }
    }
}
