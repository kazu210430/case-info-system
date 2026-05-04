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
        public void ResolveKernelWorkbook_WhenContextIsNull_ReturnsNullWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            int fallbackResolveCalls = 0;
            int fallbackFindCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
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

            Excel.Workbook resolved = service.ResolveKernelWorkbook((WorkbookContext)null);

            Assert.Null(resolved);
            Assert.Equal(0, openKernelCalls);
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
            Excel.Workbook contextKernelWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = @"C:\root"
                }
            };
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
        public void ResolveKernelWorkbook_WhenContextKernelWorkbookRootMismatches_ReturnsNullWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            int fallbackResolveCalls = 0;
            int fallbackFindCalls = 0;
            Excel.Workbook contextKernelWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = @"C:\other-root"
                }
            };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    },
                    ResolveKernelWorkbookPath = root =>
                    {
                        fallbackResolveCalls++;
                        return root + "\\kernel.xlsm";
                    },
                    FindOpenWorkbook = path =>
                    {
                        fallbackFindCalls++;
                        return new Excel.Workbook();
                    }
                });

            Excel.Workbook resolved = service.ResolveKernelWorkbook(
                new WorkbookContext(contextKernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", "kernel.xlsm", "shMasterList"));

            Assert.Null(resolved);
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
        public void CloseHomeSession_WhenValidHomeBindingExists_UsesBoundWorkbookWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            Excel.Workbook observedCloseWorkbook = null;
            Excel.Workbook boundWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\root\kernel.xlsm",
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = @"C:\root"
                }
            };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    },
                    HasOtherVisibleWorkbook = _ => false,
                    HasOtherWorkbook = _ => false,
                    CloseKernelWorkbookWithoutLifecycle = workbook => observedCloseWorkbook = workbook,
                    ReleaseHomeDisplay = showExcel => { },
                    QuitApplication = () => { }
                });

            bool bound = service.BindHomeWorkbook(
                new WorkbookContext(boundWorkbook, null, WorkbookRole.Kernel, @"C:\root", boundWorkbook.FullName, "shHOME"));

            Assert.True(bound);
            Assert.True(service.HasValidHomeWorkbookBinding());

            service.CloseHomeSession();

            Assert.Same(boundWorkbook, observedCloseWorkbook);
            Assert.Equal(0, openKernelCalls);
            Assert.False(service.HasValidHomeWorkbookBinding());
        }

        [Fact]
        public void TryGetValidHomeWorkbookBinding_WhenBindingIsValid_ReturnsWorkbookAndSystemRoot()
        {
            Excel.Workbook boundWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\root\kernel.xlsm",
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = @"C:\root"
                }
            };
            var service = CreateService(new KernelWorkbookService.KernelWorkbookServiceTestHooks());

            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(boundWorkbook, null, WorkbookRole.Kernel, @"C:\root", boundWorkbook.FullName, "shHOME")));

            Assert.True(service.TryGetValidHomeWorkbookBinding(out Excel.Workbook resolvedWorkbook, out string systemRoot));
            Assert.Same(boundWorkbook, resolvedWorkbook);
            Assert.Equal(@"C:\root", systemRoot);
        }

        [Fact]
        public void BindHomeWorkbook_WhenKernelWorkbookIsProvided_BindsUsingWorkbookSystemRootWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            Excel.Workbook boundWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\root\kernel.xlsm",
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = @"C:\root"
                }
            };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    }
                });

            Assert.True(service.BindHomeWorkbook(boundWorkbook));
            Assert.True(service.TryGetValidHomeWorkbookBinding(out Excel.Workbook resolvedWorkbook, out string systemRoot));
            Assert.Same(boundWorkbook, resolvedWorkbook);
            Assert.Equal(@"C:\root", systemRoot);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void BindHomeWorkbook_WhenWorkbookSystemRootIsMissing_FailsClosedWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            Excel.Workbook boundWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\root\kernel.xlsm",
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    }
                });

            Assert.False(service.BindHomeWorkbook(boundWorkbook));
            Assert.False(service.TryGetValidHomeWorkbookBinding(out Excel.Workbook resolvedWorkbook, out string systemRoot));
            Assert.Null(resolvedWorkbook);
            Assert.Equal(string.Empty, systemRoot);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void TryGetValidHomeWorkbookBinding_WhenBindingBecomesInvalid_ReturnsFalseWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["SYSTEM_ROOT"] = @"C:\root"
            };
            Excel.Workbook boundWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\root\kernel.xlsm",
                CustomDocumentProperties = properties
            };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    }
                });

            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(boundWorkbook, null, WorkbookRole.Kernel, @"C:\root", boundWorkbook.FullName, "shHOME")));

            properties["SYSTEM_ROOT"] = @"C:\other-root";

            Assert.False(service.TryGetValidHomeWorkbookBinding(out Excel.Workbook resolvedWorkbook, out string systemRoot));
            Assert.Null(resolvedWorkbook);
            Assert.Equal(string.Empty, systemRoot);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void LoadSettings_WhenHomeBindingBecomesInvalid_ReturnsDefaultStateWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["SYSTEM_ROOT"] = @"C:\root"
            };
            Excel.Workbook boundWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\root\kernel.xlsm",
                CustomDocumentProperties = properties
            };
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    }
                });

            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(boundWorkbook, null, WorkbookRole.Kernel, @"C:\root", boundWorkbook.FullName, "shHOME")));

            properties["SYSTEM_ROOT"] = @"C:\other-root";

            KernelSettingsState state = service.LoadSettings();

            Assert.Equal(string.Empty, state.SystemRoot);
            Assert.Equal(string.Empty, state.DefaultRoot);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void SelectAndSaveDefaultRoot_WhenHomeBindingIsMissing_FailsClosedWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        openKernelCalls++;
                        return new Excel.Workbook();
                    }
                });

            string selectedRoot = service.SelectAndSaveDefaultRoot();

            Assert.Null(selectedRoot);
            Assert.Equal(0, openKernelCalls);
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
        public void ResolveKernelWorkbook_WhenAvailabilityChanges_TransitionsFromNullToFallbackToUnavailable()
        {
            var callLog = new List<string>();
            int phase = 0;
            Excel.Workbook fallback = new Excel.Workbook();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () =>
                    {
                        callLog.Add("get:" + phase.ToString());
                        return null;
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

            Assert.Null(first);
            Assert.Same(fallback, second);
            Assert.Null(third);
            Assert.Equal(
                new[]
                {
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
