using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class CasePaneCacheRefreshNotificationServiceTests
    {
        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenReasonAndCacheAreEligible_FiresNotification()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook();
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult);

            Assert.Single(notifications);
            Assert.Equal("WorkbookOpen", notifications[0]);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenWorkbookWasDirty_RestoresDirtyState()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook
            {
                Saved = true
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult, originalSavedState: false);

            Assert.False(workbook.Saved);
            Assert.Single(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenWorkbookWasClean_RestoresCleanState()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook
            {
                Saved = false
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult, originalSavedState: true);

            Assert.True(workbook.Saved);
            Assert.Single(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenCacheUpdatedButNotificationIsSkipped_RestoresOriginalDirtyState()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook
            {
                Saved = true
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "SheetActivate", buildResult, originalSavedState: false);

            Assert.False(workbook.Saved);
            Assert.Empty(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenCacheWasNotUpdated_DoesNotRestoreSavedState()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook
            {
                Saved = true
            };
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: false);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult, originalSavedState: false);

            Assert.True(workbook.Saved);
            Assert.Empty(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenReasonIsNotEligible_DoesNotFireNotification()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook();
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "SheetActivate", buildResult);

            Assert.Empty(notifications);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenCaseListStateUpdated_DoesNotFireNotification()
        {
            var notifications = new List<string>();
            var service = CreateService(new List<string>(), reason => notifications.Add(reason));
            var workbook = new Excel.Workbook();
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            service.NotifyCasePaneUpdatedIfNeeded(workbook, "DocumentCommandService.CaseListStateUpdated", buildResult);

            Assert.Empty(notifications);
        }

        [Fact]
        public void TryGetWorkbookSavedState_WhenWorkbookIsNull_ReturnsNull()
        {
            var service = CreateService(new List<string>());

            bool? result = service.TryGetWorkbookSavedState(null);

            Assert.Null(result);
        }

        [Fact]
        public void TryGetWorkbookSavedState_WhenWorkbookExists_ReturnsSavedState()
        {
            var service = CreateService(new List<string>());
            var workbook = new Excel.Workbook
            {
                Saved = true
            };

            bool? result = service.TryGetWorkbookSavedState(workbook);

            Assert.True(result);
        }

        [Fact]
        public void NotifyCasePaneUpdatedIfNeeded_WhenNotificationCallbackThrows_LogsAndDoesNotThrow()
        {
            var logs = new List<string>();
            var service = CreateService(logs, reason => throw new InvalidOperationException("boom"));
            var workbook = new Excel.Workbook();
            var buildResult = new TaskPaneSnapshotBuilderService.TaskPaneBuildResult("snapshot", updatedCaseSnapshotCache: true);

            Exception ex = null;
            try
            {
                service.NotifyCasePaneUpdatedIfNeeded(workbook, "WorkbookOpen", buildResult);
            }
            catch (Exception caught)
            {
                ex = caught;
            }

            Assert.Null(ex);
            Assert.True(logs.Exists(message => message.IndexOf("NotifyCasePaneUpdatedIfNeeded failed.", StringComparison.Ordinal) >= 0));
        }

        private static CasePaneCacheRefreshNotificationService CreateService(List<string> logs, Action<string> notification = null)
        {
            return new CasePaneCacheRefreshNotificationService(
                OrchestrationTestSupport.CreateLogger(logs),
                workbook => workbook == null ? string.Empty : (workbook.FullName ?? string.Empty),
                notification);
        }
    }
}
