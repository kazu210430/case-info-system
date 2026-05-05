using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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
        public void LoadSettings_WhenHomeBindingIsMissing_ReturnsDefaultStateWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                });

            KernelSettingsState state = service.LoadSettings();

            Assert.Equal(string.Empty, state.SystemRoot);
            Assert.Equal(string.Empty, state.DefaultRoot);
            Assert.Equal("YYYY", state.NameRuleA);
            Assert.Equal("DOC", state.NameRuleB);
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
                });

            string selectedRoot = service.SelectAndSaveDefaultRoot();

            Assert.Null(selectedRoot);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void ResolveWorkbookForHomeDisplayOrClose_WhenHomeBindingIsMissing_ReturnsNullWithoutFallbackLookup()
        {
            int openKernelCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                });

            Excel.Workbook resolved = (Excel.Workbook)typeof(KernelWorkbookService)
                .GetMethod("ResolveWorkbookForHomeDisplayOrClose", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(service, new object[] { "test" });

            Assert.Null(resolved);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void ShowKernelWorkbookWindows_WhenHomeBindingIsMissing_DoesNotLookupOpenKernelWorkbook()
        {
            int openKernelCalls = 0;
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                });

            typeof(KernelWorkbookService)
                .GetMethod("ShowKernelWorkbookWindows", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(service, new object[] { true });

            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void CloseHomeSession_WhenNoOtherWorkbookExists_QuitsWithoutRestore()
        {
            int openKernelCalls = 0;
            int quitCalls = 0;
            var releaseCalls = new List<bool>();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    HasOtherVisibleWorkbook = _ => false,
                    HasOtherWorkbook = _ => false,
                    ReleaseHomeDisplay = showExcel => releaseCalls.Add(showExcel),
                    QuitApplication = () => quitCalls++
                });

            service.CloseHomeSession();

            Assert.Single(releaseCalls);
            Assert.False(releaseCalls[0]);
            Assert.Equal(1, quitCalls);
            Assert.Equal(0, openKernelCalls);
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
                Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
                var service = new KernelWorkbookService(
                    interactionState,
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    new KernelWorkbookService.KernelWorkbookServiceTestHooks
                    {
                        HasOtherVisibleWorkbook = _ => true,
                        HasOtherWorkbook = _ => true,
                        ApplyHomeDisplayVisibility = () => { },
                        ConcealKernelWorkbookWindowsForCaseCreationClose = workbook => callLog.Add("conceal"),
                        SaveAndCloseKernelWorkbook = workbook => callLog.Add("save-close"),
                        ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                        DismissPreparedHomeDisplayState = reason => callLog.Add("dismiss"),
                        QuitApplication = () => quitCalls++
                    });

                Assert.True(service.BindHomeWorkbook(
                    new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));
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
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
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
            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));
            service.CloseHomeSession();

            Assert.Equal(0, quitCalls);
            Assert.Equal(new[] { "request-close" }, callLog);
        }

        [Fact]
        public void CloseHomeSession_WhenManagedCloseIsRetriedAfterRejection_ReusesPreparedState()
        {
            var callLog = new List<string>();
            int requestCalls = 0;
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            var lifecycleService = new KernelWorkbookLifecycleService();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => callLog.Add("apply"),
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

            service.SetLifecycleService(lifecycleService);
            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));

            service.PrepareForHomeDisplay();
            service.CloseHomeSession();
            service.CloseHomeSession();
            Assert.True(service.HasValidHomeWorkbookBinding());
            lifecycleService.SimulateManagedCloseSuccess(kernelWorkbook);

            Assert.Equal(
                new[]
                {
                    "apply",
                    "request:1",
                    "request:2",
                    "release:True"
                },
                callLog);
            Assert.False(service.HasValidHomeWorkbookBinding());
        }

        [Fact]
        public void CloseHomeSession_WhenManagedCloseFails_PreservesBindingUntilRetrySucceeds()
        {
            var callLog = new List<string>();
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            var lifecycleService = new KernelWorkbookLifecycleService();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => callLog.Add("apply"),
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => callLog.Add("quit")
                });

            service.SetLifecycleService(lifecycleService);
            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));

            service.PrepareForHomeDisplay();
            service.CloseHomeSession();
            lifecycleService.SimulateManagedCloseFailure(
                kernelWorkbook,
                new COMException("managed close failed", unchecked((int)0x80020005)));

            Assert.Equal(new[] { "apply" }, callLog);
            Assert.True(service.HasValidHomeWorkbookBinding());

            service.CloseHomeSession();
            lifecycleService.SimulateManagedCloseSuccess(kernelWorkbook);

            Assert.Equal(
                new[]
                {
                    "apply",
                    "release:True"
                },
                callLog);
            Assert.False(service.HasValidHomeWorkbookBinding());
        }

        [Fact]
        public void RequestCloseHomeSessionFromForm_WhenManagedCloseSucceeds_DefersReleaseUntilFinalize()
        {
            int readyNotifications = 0;
            int failureNotifications = 0;
            var callLog = new List<string>();
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            var lifecycleService = new KernelWorkbookLifecycleService();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => { },
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => callLog.Add("quit")
                });

            service.SetLifecycleService(lifecycleService);
            service.RegisterHomeSessionCloseObserver(
                () => readyNotifications++,
                () => failureNotifications++);
            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));
            service.PrepareForHomeDisplay();

            KernelHomeSessionCloseRequestStatus status = service.RequestCloseHomeSessionFromForm(false, "test");

            Assert.Equal(KernelHomeSessionCloseRequestStatus.Pending, status);
            Assert.Empty(callLog);
            Assert.True(service.HasValidHomeWorkbookBinding());

            lifecycleService.SimulateManagedCloseSuccess(kernelWorkbook);

            Assert.Equal(1, readyNotifications);
            Assert.Equal(0, failureNotifications);
            Assert.Empty(callLog);
            Assert.True(service.HasValidHomeWorkbookBinding());

            service.FinalizePendingHomeSessionCloseAfterFormClosed();

            Assert.Equal(new[] { "release:True" }, callLog);
            Assert.False(service.HasValidHomeWorkbookBinding());
        }

        [Fact]
        public void RequestCloseHomeSessionFromForm_WhenManagedCloseFails_NotifiesFailureWithoutRelease()
        {
            int readyNotifications = 0;
            int failureNotifications = 0;
            var callLog = new List<string>();
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            var lifecycleService = new KernelWorkbookLifecycleService();
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => { },
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => callLog.Add("quit")
                });

            service.SetLifecycleService(lifecycleService);
            service.RegisterHomeSessionCloseObserver(
                () => readyNotifications++,
                () => failureNotifications++);
            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));
            service.PrepareForHomeDisplay();

            KernelHomeSessionCloseRequestStatus status = service.RequestCloseHomeSessionFromForm(false, "test");

            Assert.Equal(KernelHomeSessionCloseRequestStatus.Pending, status);

            lifecycleService.SimulateManagedCloseFailure(
                kernelWorkbook,
                new COMException("managed close failed", unchecked((int)0x80020005)));

            Assert.Equal(0, readyNotifications);
            Assert.Equal(1, failureNotifications);
            Assert.Empty(callLog);
            Assert.True(service.HasValidHomeWorkbookBinding());

            service.FinalizePendingHomeSessionCloseAfterFormClosed();

            Assert.Empty(callLog);
            Assert.True(service.HasValidHomeWorkbookBinding());
        }

        [Fact]
        public void RequestCloseHomeSessionFromForm_WhenSaveCloseSucceeds_DefersReleaseUntilFinalize()
        {
            var callLog = new List<string>();
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            var service = CreateService(
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    ApplyHomeDisplayVisibility = () => { },
                    HasOtherVisibleWorkbook = _ => true,
                    HasOtherWorkbook = _ => true,
                    SaveAndCloseKernelWorkbook = workbook => callLog.Add("save-close"),
                    ReleaseHomeDisplay = showExcel => callLog.Add("release:" + showExcel.ToString()),
                    QuitApplication = () => callLog.Add("quit")
                });

            Assert.True(service.BindHomeWorkbook(
                new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, @"C:\root", kernelWorkbook.FullName, "shHOME")));
            service.PrepareForHomeDisplay();

            KernelHomeSessionCloseRequestStatus status = service.RequestCloseHomeSessionFromForm(true, "test");

            Assert.Equal(KernelHomeSessionCloseRequestStatus.Completed, status);
            Assert.Equal(new[] { "save-close" }, callLog);
            Assert.True(service.HasValidHomeWorkbookBinding());

            service.FinalizePendingHomeSessionCloseAfterFormClosed();

            Assert.Equal(
                new[]
                {
                    "save-close",
                    "release:True"
                },
                callLog);
            Assert.False(service.HasValidHomeWorkbookBinding());
        }

        [Fact]
        public void TryShowSheetByCodeName_WhenHomeDisplayPrepared_RestoresWorkbookWindowVisibilityBeforeSheetActivation()
        {
            var loggerMessages = new List<string>();
            var interactionMessages = new List<string>();
            string systemRoot = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(systemRoot);
            string kernelWorkbookPath = Path.Combine(systemRoot, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
            File.WriteAllText(kernelWorkbookPath, string.Empty);
            try
            {
                var application = new Excel.Application
                {
                    Visible = true,
                    Hwnd = 1
                };
                var logger = OrchestrationTestSupport.CreateLogger(loggerMessages);
                var excelInteropService = new ExcelInteropService(application, logger, new PathCompatibilityService());
                var excelWindowRecoveryService = new ExcelWindowRecoveryService(application, excelInteropService, logger);
                var service = new KernelWorkbookService(
                    application,
                    excelInteropService,
                    excelWindowRecoveryService,
                    OrchestrationTestSupport.CreateKernelCaseInteractionState(interactionMessages),
                    logger);
                Excel.Workbook kernelWorkbook = CreateKernelWorkbook(systemRoot);
                kernelWorkbook.FullName = kernelWorkbookPath;
                kernelWorkbook.Path = systemRoot;
                var window = new Excel.Window
                {
                    Visible = true,
                    WindowState = Excel.XlWindowState.xlNormal
                };
                var worksheet = new Excel.Worksheet
                {
                    CodeName = "shUserData",
                    Name = "UserData",
                    Parent = kernelWorkbook
                };
                kernelWorkbook.Windows.Add(window);
                kernelWorkbook.Worksheets.Add(worksheet);
                kernelWorkbook.ActiveSheet = worksheet;
                application.Workbooks.Add(kernelWorkbook);
                kernelWorkbook.Activate();

                var context = new WorkbookContext(kernelWorkbook, null, WorkbookRole.Kernel, systemRoot, kernelWorkbook.FullName, "shHOME");
                Assert.True(service.BindHomeWorkbook(context));

                service.PrepareForHomeDisplayFromSheet();
                Assert.False(window.Visible);

                bool shown = service.TryShowSheetByCodeName(context, "shUserData", "test");

                Assert.True(shown);
                Assert.True(window.Visible);
                Assert.Equal(Excel.XlWindowState.xlNormal, window.WindowState);
                Assert.Same(kernelWorkbook, application.ActiveWorkbook);
            }
            finally
            {
                try
                {
                    if (Directory.Exists(systemRoot))
                    {
                        Directory.Delete(systemRoot, recursive: true);
                    }
                }
                catch
                {
                }
            }
        }

        [Fact]
        public void CloseKernelWorkbookWithoutLifecycleCore_UsesInteropHelperOptionalArgumentsAndRestoresDisplayAlerts()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var service = CreateRealService(application);
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            application.Workbooks.Add(kernelWorkbook);

            InvokePrivate(service, "CloseKernelWorkbookWithoutLifecycleCore", kernelWorkbook);

            Assert.Equal(1, kernelWorkbook.CloseCallCount);
            Assert.False(kernelWorkbook.LastCloseSaveChanges.GetValueOrDefault());
            Assert.Same(Type.Missing, kernelWorkbook.LastCloseFilename);
            Assert.Same(Type.Missing, kernelWorkbook.LastCloseRouteWorkbook);
            Assert.True(application.DisplayAlerts);
        }

        [Fact]
        public void CloseKernelWorkbookWithoutLifecycleCore_WhenCloseThrows_RestoresDisplayAlerts()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var service = CreateRealService(application);
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            kernelWorkbook.CloseBehavior = () => throw new InvalidOperationException("close failed");
            application.Workbooks.Add(kernelWorkbook);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => InvokePrivate(service, "CloseKernelWorkbookWithoutLifecycleCore", kernelWorkbook));

            Assert.Equal("close failed", exception.Message);
            Assert.True(application.DisplayAlerts);
        }

        [Fact]
        public void SaveAndCloseKernelWorkbook_UsesInteropHelperOptionalArgumentsAndRestoresDisplayAlerts()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var service = CreateRealService(application);
            Excel.Workbook kernelWorkbook = CreateKernelWorkbook(@"C:\root");
            application.Workbooks.Add(kernelWorkbook);

            InvokePrivate(service, "SaveAndCloseKernelWorkbook", kernelWorkbook);

            Assert.Equal(1, kernelWorkbook.SaveCallCount);
            Assert.Equal(1, kernelWorkbook.CloseCallCount);
            Assert.False(kernelWorkbook.LastCloseSaveChanges.GetValueOrDefault());
            Assert.Same(Type.Missing, kernelWorkbook.LastCloseFilename);
            Assert.Same(Type.Missing, kernelWorkbook.LastCloseRouteWorkbook);
            Assert.True(application.DisplayAlerts);
        }

        [Fact]
        public void QuitApplicationCore_WhenQuitSucceeds_DoesNotRestoreDisplayAlertsAfterSuccessfulQuit()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var service = CreateRealService(application);

            InvokePrivate(service, "QuitApplicationCore");

            Assert.Equal(1, application.QuitCallCount);
            Assert.False(application.DisplayAlerts);
        }

        [Fact]
        public void QuitApplicationCore_WhenQuitThrows_RestoresDisplayAlerts()
        {
            var application = new Excel.Application
            {
                DisplayAlerts = true,
                QuitBehavior = () => throw new InvalidOperationException("quit failed")
            };
            var service = CreateRealService(application);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => InvokePrivate(service, "QuitApplicationCore"));

            Assert.Equal("quit failed", exception.Message);
            Assert.Equal(1, application.QuitCallCount);
            Assert.True(application.DisplayAlerts);
        }

        private static KernelWorkbookService CreateService(KernelWorkbookService.KernelWorkbookServiceTestHooks hooks)
        {
            return new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                hooks);
        }

        private static KernelWorkbookService CreateRealService(Excel.Application application)
        {
            var loggerMessages = new List<string>();
            var interactionMessages = new List<string>();
            var logger = OrchestrationTestSupport.CreateLogger(loggerMessages);
            var excelInteropService = new ExcelInteropService(application, logger, new PathCompatibilityService());
            var excelWindowRecoveryService = new ExcelWindowRecoveryService(application, excelInteropService, logger);

            return new KernelWorkbookService(
                application,
                excelInteropService,
                excelWindowRecoveryService,
                OrchestrationTestSupport.CreateKernelCaseInteractionState(interactionMessages),
                logger);
        }

        private static void InvokePrivate(object target, string methodName, params object[] args)
        {
            MethodInfo method = target.GetType().GetMethod(methodName, BindingFlags.Instance | BindingFlags.NonPublic);

            try
            {
                method.Invoke(target, args);
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
        }

        private static Excel.Workbook CreateKernelWorkbook(string systemRoot)
        {
            return new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = systemRoot + @"\kernel.xlsm",
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = systemRoot
                }
            };
        }
    }
}
