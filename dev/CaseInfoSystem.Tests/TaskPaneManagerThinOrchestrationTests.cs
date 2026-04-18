using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneManagerThinOrchestrationTests
    {
        [Fact]
        public void RefreshPane_WhenContextIsInvalid_DoesNotInvokeShowAndHidesExistingHosts()
        {
            var logs = new List<string>();
            var hidden = new List<string>();
            int showAttempts = 0;
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(logs),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnHideHost = (windowKey, reason) => hidden.Add(windowKey),
                    TryShowHost = (windowKey, reason) =>
                    {
                        showAttempts++;
                        return true;
                    }
                });

            manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "case"));
            manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "kernel"));

            bool refreshed = manager.RefreshPane(null, "WorkbookOpen");

            Assert.False(refreshed);
            Assert.Equal(0, showAttempts);
            Assert.Equal(2, hidden.Count);
            Assert.Contains("case", hidden);
            Assert.Contains("kernel", hidden);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenReasonAndCacheAreEligible_FiresNotification()
        {
            var notifications = new List<string>();
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifications.Add(reason)
                });

            var workbook = new Excel.Workbook();
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            manager.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult);

            Assert.Single(notifications);
            Assert.Equal("WorkbookOpen", notifications[0]);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenWorkbookWasDirty_RestoresDirtyState()
        {
            var notifications = new List<string>();
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifications.Add(reason)
                });

            var workbook = new Excel.Workbook
            {
                Saved = true
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            manager.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult, originalSavedState: false);

            Assert.False(workbook.Saved);
            Assert.Single(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenWorkbookWasClean_RestoresCleanState()
        {
            var notifications = new List<string>();
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifications.Add(reason)
                });

            var workbook = new Excel.Workbook
            {
                Saved = false
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            manager.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult, originalSavedState: true);

            Assert.True(workbook.Saved);
            Assert.Single(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenCacheUpdatedButNotificationIsSkipped_RestoresOriginalDirtyState()
        {
            var notifications = new List<string>();
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifications.Add(reason)
                });

            var workbook = new Excel.Workbook
            {
                Saved = true
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            manager.NotifyCasePaneUpdatedIfNeeded(workbook, "SheetActivate", buildResult, originalSavedState: false);

            Assert.False(workbook.Saved);
            Assert.Empty(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenCacheWasNotUpdated_DoesNotRestoreSavedState()
        {
            var notifications = new List<string>();
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifications.Add(reason)
                });

            var workbook = new Excel.Workbook
            {
                Saved = true
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: false);

            manager.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult, originalSavedState: false);

            Assert.True(workbook.Saved);
            Assert.Empty(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenReasonIsNotEligible_DoesNotFireNotification()
        {
            var notifications = new List<string>();
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifications.Add(reason)
                });

            var workbook = new Excel.Workbook();
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            manager.NotifyCasePaneUpdatedIfNeeded(workbook, "SheetActivate", buildResult);

            Assert.Empty(notifications);
        }

        [Fact]
        public void PrepareHostsBeforeShow_WhenCaseCreationFlowUsesCaseHost_HidesOnlyNonCaseHosts()
        {
            var hidden = new List<string>();
            var stateLogs = new List<string>();
            KernelCaseInteractionState interactionState = OrchestrationTestSupport.CreateKernelCaseInteractionState(stateLogs);
            using (interactionState.BeginKernelCaseCreationFlow("test"))
            {
                var manager = new TaskPaneManager(
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    interactionState,
                    new TaskPaneManager.TaskPaneManagerTestHooks
                    {
                        OnHideHost = (windowKey, reason) => hidden.Add(windowKey)
                    });

                TaskPaneHost activeCase = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "active-case");
                manager.RegisterHost(activeCase);
                manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "other-case"));
                manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "kernel"));

                manager.PrepareHostsBeforeShow(activeCase);
            }

            Assert.Single(hidden);
            Assert.Equal("kernel", hidden[0]);
        }

        [Fact]
        public void PrepareHostsBeforeShow_WhenCaseCreationUsesNonCaseHost_HidesAllOtherHosts()
        {
            var hidden = new List<string>();
            var stateLogs = new List<string>();
            KernelCaseInteractionState interactionState = OrchestrationTestSupport.CreateKernelCaseInteractionState(stateLogs);
            using (interactionState.BeginKernelCaseCreationFlow("test"))
            {
                var manager = new TaskPaneManager(
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    interactionState,
                    new TaskPaneManager.TaskPaneManagerTestHooks
                    {
                        OnHideHost = (windowKey, reason) => hidden.Add(windowKey)
                    });

                TaskPaneHost activeKernel = OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "active-kernel");
                manager.RegisterHost(activeKernel);
                manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "case"));
                manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new AccountingNavigationControl(), "accounting"));

                manager.PrepareHostsBeforeShow(activeKernel);
            }

            Assert.Equal(2, hidden.Count);
            Assert.Contains("case", hidden);
            Assert.Contains("accounting", hidden);
        }

        [Fact]
        public void PrepareHostsBeforeShow_WhenCaseCreationFlowEnds_ReturnsToNormalHideAllBehavior()
        {
            var hidden = new List<string>();
            var stateLogs = new List<string>();
            KernelCaseInteractionState interactionState = OrchestrationTestSupport.CreateKernelCaseInteractionState(stateLogs);
            var manager = new TaskPaneManager(
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                interactionState,
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnHideHost = (windowKey, reason) => hidden.Add(windowKey)
                });

            TaskPaneHost activeCase = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "active-case");
            manager.RegisterHost(activeCase);
            manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "other-case"));
            manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "kernel"));

            using (interactionState.BeginKernelCaseCreationFlow("test"))
            {
                manager.PrepareHostsBeforeShow(activeCase);
            }

            Assert.Equal(new[] { "kernel" }, hidden);

            hidden.Clear();

            manager.PrepareHostsBeforeShow(activeCase);

            Assert.Empty(hidden);
        }

        [Fact]
        public void RefreshPane_WhenCasePaneRefreshesDuringCaseCreation_InvokesNotificationBeforeHideAndShow()
        {
            var callLog = new List<string>();
            var lifecycleLogs = new List<string>();
            KernelCaseInteractionState interactionState = OrchestrationTestSupport.CreateKernelCaseInteractionState(lifecycleLogs);
            var builder = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook =>
                {
                    callLog.Add("build");
                    return new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: true);
                }
            };
            var manager = CreateFullManager(
                interactionState,
                builder,
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => callLog.Add("notify:" + reason),
                    OnHideHost = (windowKey, reason) => callLog.Add("hide:" + windowKey),
                    TryShowHost = (windowKey, reason) =>
                    {
                        callLog.Add("show:" + windowKey);
                        return true;
                    }
                });

            manager.RegisterHost(OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "other-kernel"));

            using (interactionState.BeginKernelCaseCreationFlow("test"))
            {
                bool refreshed = manager.RefreshPane(
                    new WorkbookContext(
                        new Excel.Workbook { FullName = @"C:\cases\case.xlsx", Name = "case.xlsx" },
                        new Excel.Window { Hwnd = 101 },
                        WorkbookRole.Case,
                        @"C:\cases",
                        @"C:\cases\case.xlsx",
                        "shHOME"),
                    "WorkbookOpen");

                Assert.True(refreshed);
            }

            Assert.Equal(
                new[]
                {
                    "build",
                    "notify:WorkbookOpen",
                    "hide:other-kernel",
                    "show:101"
                },
                callLog);
        }

        [Fact]
        public void RefreshPane_WhenSignatureIsUnchanged_SkipsSnapshotRebuildOnSecondRefresh()
        {
            int buildCalls = 0;
            int showCalls = 0;
            var builder = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook =>
                {
                    buildCalls++;
                    return new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: false);
                }
            };
            var manager = CreateFullManager(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                builder,
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    TryShowHost = (windowKey, reason) =>
                    {
                        showCalls++;
                        return true;
                    }
                });
            var context = new WorkbookContext(
                new Excel.Workbook { FullName = @"C:\cases\case.xlsx", Name = "case.xlsx" },
                new Excel.Window { Hwnd = 202 },
                WorkbookRole.Case,
                @"C:\cases",
                @"C:\cases\case.xlsx",
                "shHOME");

            bool first = manager.RefreshPane(context, "WindowActivate");
            bool second = manager.RefreshPane(context, "WindowActivate");

            Assert.True(first);
            Assert.True(second);
            Assert.Equal(1, buildCalls);
            Assert.Equal(2, showCalls);
        }

        [Fact]
        public void RefreshPane_WhenShowFailsAfterRender_RetryShowsWithoutRebuildOrRenotify()
        {
            int buildCalls = 0;
            int notifyCalls = 0;
            int showCalls = 0;
            var builder = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook =>
                {
                    buildCalls++;
                    return new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: true);
                }
            };
            var manager = CreateFullManager(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                builder,
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifyCalls++,
                    TryShowHost = (windowKey, reason) =>
                    {
                        showCalls++;
                        return showCalls > 1;
                    }
                });
            var context = new WorkbookContext(
                new Excel.Workbook { FullName = @"C:\cases\case.xlsx", Name = "case.xlsx" },
                new Excel.Window { Hwnd = 404 },
                WorkbookRole.Case,
                @"C:\cases",
                @"C:\cases\case.xlsx",
                "shHOME");

            bool first = manager.RefreshPane(context, "WorkbookOpen");
            bool second = manager.RefreshPane(context, "WorkbookOpen");

            Assert.False(first);
            Assert.True(second);
            Assert.Equal(1, buildCalls);
            Assert.Equal(1, notifyCalls);
            Assert.Equal(2, showCalls);
        }

        [Fact]
        public void RefreshPane_WhenBuildThrows_DoesNotNotifyOrShowAndRetriesBuildOnNextRefresh()
        {
            int buildCalls = 0;
            int notifyCalls = 0;
            int showCalls = 0;
            var builder = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook =>
                {
                    buildCalls++;
                    if (buildCalls == 1)
                    {
                        throw new InvalidOperationException("build failed");
                    }

                    return new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: true);
                }
            };
            var manager = CreateFullManager(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                builder,
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifyCalls++,
                    TryShowHost = (windowKey, reason) =>
                    {
                        showCalls++;
                        return true;
                    }
                });
            var context = new WorkbookContext(
                new Excel.Workbook { FullName = @"C:\cases\case.xlsx", Name = "case.xlsx" },
                new Excel.Window { Hwnd = 505 },
                WorkbookRole.Case,
                @"C:\cases",
                @"C:\cases\case.xlsx",
                "shHOME");

            Assert.Throws<InvalidOperationException>(() => manager.RefreshPane(context, "WorkbookOpen"));

            bool retried = manager.RefreshPane(context, "WorkbookOpen");

            Assert.True(retried);
            Assert.Equal(2, buildCalls);
            Assert.Equal(1, notifyCalls);
            Assert.Equal(1, showCalls);
        }

        [Fact]
        public void RefreshPane_WhenShowFailsAfterReuse_DoesNotRebuildOrNotify()
        {
            int buildCalls = 0;
            int notifyCalls = 0;
            int showCalls = 0;
            var builder = new TaskPaneSnapshotBuilderService
            {
                OnBuildSnapshotText = workbook =>
                {
                    buildCalls++;
                    return new TaskPaneSnapshotBuilderService.TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: true);
                }
            };
            var manager = CreateFullManager(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                builder,
                new TaskPaneManager.TaskPaneManagerTestHooks
                {
                    OnCasePaneUpdatedNotification = reason => notifyCalls++,
                    TryShowHost = (windowKey, reason) =>
                    {
                        showCalls++;
                        return showCalls == 1;
                    }
                });
            var context = new WorkbookContext(
                new Excel.Workbook { FullName = @"C:\cases\case.xlsx", Name = "case.xlsx" },
                new Excel.Window { Hwnd = 303 },
                WorkbookRole.Case,
                @"C:\cases",
                @"C:\cases\case.xlsx",
                "shHOME");

            bool first = manager.RefreshPane(context, "WindowActivate");
            bool second = manager.RefreshPane(context, "WindowActivate");

            Assert.True(first);
            Assert.False(second);
            Assert.Equal(1, buildCalls);
            Assert.Equal(0, notifyCalls);
            Assert.Equal(2, showCalls);
        }

        private static TaskPaneManager CreateFullManager(
            KernelCaseInteractionState interactionState,
            TaskPaneSnapshotBuilderService snapshotBuilderService,
            TaskPaneManager.TaskPaneManagerTestHooks hooks)
        {
            return new TaskPaneManager(
                new CaseInfoSystem.ExcelAddIn.ThisAddIn(),
                new ExcelInteropService(),
                snapshotBuilderService ?? new TaskPaneSnapshotBuilderService(),
                CreateDocumentCommandService(),
                new DocumentEligibilityDiagnosticsService(),
                new DocumentMasterCatalogDiagnosticsService(),
                new DocumentNamePromptService(),
                new KernelCommandService(),
                new AccountingSheetCommandService(),
                new CaseTaskPaneViewStateBuilder(),
                new AccountingInternalCommandService(),
                interactionState,
                new UserErrorService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                hooks);
        }

        private static DocumentCommandService CreateDocumentCommandService()
        {
            return new DocumentCommandService(
                new CaseInfoSystem.ExcelAddIn.ThisAddIn(),
                new InlineScreenUpdatingExecutionBridge(),
                new NoOpTaskPaneRefreshSuppressionBridge(),
                new CollectingActiveTaskPaneRefreshBridge(),
                new DocumentExecutionModeService(OrchestrationTestSupport.CreateLogger(new List<string>()), new ExcelInteropService()),
                new DocumentExecutionEligibilityService(),
                new DocumentExecutionPolicyService(),
                new DocumentCreateService(),
                new AccountingSetCommandService(),
                new CaseListRegistrationService(),
                new CaseContextFactory(),
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()));
        }

        private sealed class InlineScreenUpdatingExecutionBridge : IScreenUpdatingExecutionBridge
        {
            public void Execute(Action action)
            {
                action?.Invoke();
            }
        }

        private sealed class NoOpTaskPaneRefreshSuppressionBridge : ITaskPaneRefreshSuppressionBridge
        {
            public IDisposable Enter(string reason)
            {
                return new NoOpDisposable();
            }
        }

        private sealed class CollectingActiveTaskPaneRefreshBridge : IActiveTaskPaneRefreshBridge
        {
            public void RequestRefresh(string reason)
            {
            }
        }

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
