using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class PostCloseFollowUpSchedulerTests
    {
        [Fact]
        public void QuitExcelIfNoVisibleWorkbook_WhenNoVisibleWorkbook_Remains_DoesNotRestoreDisplayAlertsAfterSuccessfulQuit()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var hiddenWorkbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                Path = @"C:\cases"
            };

            AddWorkbookWindow(hiddenWorkbook, visible: false);
            application.Workbooks.Add(hiddenWorkbook);

            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeQuitExcelIfNoVisibleWorkbook(scheduler);

            Assert.Equal(1, application.QuitCallCount);
            Assert.False(application.DisplayAlerts);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionCompleted")
                    && ContainsFragment(message, "hasVisibleWorkbook=False")
                    && ContainsFragment(message, "quitAttempted=True")
                    && ContainsFragment(message, "quitCompleted=True")
                    && ContainsFragment(message, "outcomeReason=noVisibleWorkbookQuitCompleted")
                    && ContainsFragment(message, "targetWorkbookStillOpen=unknown"));
        }

        [Fact]
        public void QuitExcelIfNoVisibleWorkbook_WhenQuitThrows_RestoresDisplayAlerts()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application
            {
                DisplayAlerts = true,
                QuitBehavior = () => throw new InvalidOperationException("quit failed")
            };
            var hiddenWorkbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                Path = @"C:\cases"
            };

            AddWorkbookWindow(hiddenWorkbook, visible: false);
            application.Workbooks.Add(hiddenWorkbook);

            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => InvokeQuitExcelIfNoVisibleWorkbook(scheduler));

            Assert.Equal("quit failed", exception.Message);
            Assert.Equal(1, application.QuitCallCount);
            Assert.True(application.DisplayAlerts);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionFailed")
                    && ContainsFragment(message, "hasVisibleWorkbook=False")
                    && ContainsFragment(message, "quitAttempted=True")
                    && ContainsFragment(message, "quitCompleted=False")
                    && ContainsFragment(message, "outcomeReason=quitFailed")
                    && ContainsFragment(message, "targetWorkbookStillOpen=unknown"));
        }

        [Fact]
        public void QuitExcelIfNoVisibleWorkbook_WhenVisibleWorkbookRemains_DoesNotQuitExcel()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var visibleWorkbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                Path = @"C:\cases"
            };

            AddWorkbookWindow(visibleWorkbook, visible: true);
            application.Workbooks.Add(visibleWorkbook);

            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeQuitExcelIfNoVisibleWorkbook(scheduler);

            Assert.Equal(0, application.QuitCallCount);
            Assert.True(application.DisplayAlerts);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionNotRequired")
                    && ContainsFragment(message, "hasVisibleWorkbook=True")
                    && ContainsFragment(message, "quitAttempted=False")
                    && ContainsFragment(message, "quitCompleted=False")
                    && ContainsFragment(message, "outcomeReason=visibleWorkbookExists")
                    && ContainsFragment(message, "targetWorkbookStillOpen=unknown"));
            Assert.False(loggerMessages.Exists(message => ContainsFragment(message, "WhiteExcelPreventionSkipped")));
        }

        [Fact]
        public void Schedule_LogsQueuedDiagnosticWithCapturedFacts()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application();
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeSchedule(scheduler, @"C:\cases\case.xlsx", @"C:\cases");

            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionQueued")
                    && ContainsFragment(message, @"workbook=C:\cases\case.xlsx")
                    && ContainsFragment(message, "hasVisibleWorkbook=unknown")
                    && ContainsFragment(message, "quitAttempted=False")
                    && ContainsFragment(message, "quitCompleted=False")
                    && ContainsFragment(message, "outcomeReason=postCloseFollowUpQueued")
                    && ContainsFragment(message, "pendingQueueCount=1")
                    && ContainsFragment(message, "attemptsRemaining=20")
                    && ContainsFragment(message, "attemptNumber=1")
                    && ContainsFragment(message, "targetStillOpenRetriesRemaining=5")
                    && ContainsFragment(message, "excelBusyRetriesRemaining=20")
                    && ContainsFragment(message, "folderPathPresent=True")
                    && ContainsFragment(message, "targetWorkbookStillOpen=unknown"));
        }

        [Fact]
        public void ExecutePendingPostCloseQueue_WhenTargetWorkbookStillOpen_RetriesInsteadOfTerminalNotRequired()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application();
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                Path = @"C:\cases"
            };

            application.Workbooks.Add(workbook);
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeSchedule(scheduler, @"C:\cases\case.xlsx", @"C:\cases");
            InvokeExecutePendingPostCloseQueue(scheduler);

            Assert.Equal(0, application.QuitCallCount);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-request-dequeued")
                    && ContainsFragment(message, @"workbook=C:\cases\case.xlsx")
                    && ContainsFragment(message, "pendingQueueCount=0")
                    && ContainsFragment(message, "attemptsRemaining=20")
                    && ContainsFragment(message, "attemptNumber=1")
                    && ContainsFragment(message, "targetStillOpenRetriesRemaining=5")
                    && ContainsFragment(message, "workbooksCount=1")
                    && ContainsFragment(message, "activeWorkbookPresent=False"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-decision")
                    && ContainsFragment(message, "targetWorkbookStillOpen=True")
                    && ContainsFragment(message, "decision=retry-target-still-open")
                    && ContainsFragment(message, "attemptNumber=1"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-retry-scheduled")
                    && ContainsFragment(message, "retryReason=targetWorkbookStillOpen")
                    && ContainsFragment(message, "retryDelayMs=250")
                    && ContainsFragment(message, "targetWorkbookStillOpen=True")
                    && ContainsFragment(message, "attemptNumber=1")
                    && ContainsFragment(message, "nextAttemptNumber=2")
                    && ContainsFragment(message, "targetStillOpenRetriesRemaining=4"));
            Assert.DoesNotContain(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionNotRequired")
                    && ContainsFragment(message, "outcomeReason=targetWorkbookStillOpen"));
        }

        [Fact]
        public void ExecutePendingPostCloseQueue_WhenTargetClosesAfterRetryAndNoVisibleWorkbook_Remains_QuitsExcel()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                Path = @"C:\cases"
            };

            AddWorkbookWindow(workbook, visible: false);
            application.Workbooks.Add(workbook);
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeSchedule(scheduler, @"C:\cases\case.xlsx", @"C:\cases");
            InvokeExecutePendingPostCloseQueue(scheduler);

            application.Workbooks.Remove(workbook);
            InvokeExecutePendingPostCloseQueue(scheduler);

            Assert.Equal(1, application.QuitCallCount);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-decision")
                    && ContainsFragment(message, @"workbook=C:\cases\case.xlsx")
                    && ContainsFragment(message, "targetWorkbookStillOpen=False")
                    && ContainsFragment(message, "decision=scan-visible-workbooks")
                    && ContainsFragment(message, "attemptNumber=2")
                    && ContainsFragment(message, "workbooksCount=0")
                    && ContainsFragment(message, "activeWorkbookPresent=False"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionCompleted")
                    && ContainsFragment(message, @"workbook=C:\cases\case.xlsx")
                    && ContainsFragment(message, "hasVisibleWorkbook=False")
                    && ContainsFragment(message, "quitAttempted=True")
                    && ContainsFragment(message, "quitCompleted=True")
                    && ContainsFragment(message, "outcomeReason=noVisibleWorkbookQuitCompleted")
                    && ContainsFragment(message, "attemptNumber=2")
                    && ContainsFragment(message, "targetStillOpenRetriesRemaining=4"));
        }

        [Fact]
        public void ExecutePendingPostCloseQueue_WhenTargetStillOpenAfterRetryLimit_DoesNotQuitExcel()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application();
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                Path = @"C:\cases"
            };

            application.Workbooks.Add(workbook);
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeSchedule(scheduler, @"C:\cases\case.xlsx", @"C:\cases");
            for (int i = 0; i < 6; i++)
            {
                InvokeExecutePendingPostCloseQueue(scheduler);
            }

            Assert.Equal(0, application.QuitCallCount);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-decision")
                    && ContainsFragment(message, "targetWorkbookStillOpen=True")
                    && ContainsFragment(message, "decision=skip-still-open-retry-exhausted")
                    && ContainsFragment(message, "attemptNumber=6")
                    && ContainsFragment(message, "targetStillOpenRetriesRemaining=0"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionNotRequired")
                    && ContainsFragment(message, "hasVisibleWorkbook=unknown")
                    && ContainsFragment(message, "quitAttempted=False")
                    && ContainsFragment(message, "quitCompleted=False")
                    && ContainsFragment(message, "outcomeReason=targetWorkbookStillOpenRetryExhausted")
                    && ContainsFragment(message, "targetWorkbookStillOpen=True")
                    && ContainsFragment(message, "attemptNumber=6")
                    && ContainsFragment(message, "targetStillOpenRetriesRemaining=0"));
            Assert.Equal(
                5,
                loggerMessages.FindAll(message => ContainsFragment(message, "action=post-close-follow-up-retry-scheduled")
                    && ContainsFragment(message, "retryReason=targetWorkbookStillOpen")).Count);
        }

        [Fact]
        public void ExecutePendingPostCloseQueue_WhenTargetClosedButVisibleWorkbookExists_LogsVisibleWorkbookFactSeparatelyFromStillOpenFact()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application();
            var visibleWorkbook = new Excel.Workbook
            {
                FullName = @"C:\cases\other.xlsx",
                Name = "other.xlsx",
                Path = @"C:\cases"
            };

            AddWorkbookWindow(visibleWorkbook, visible: true);
            application.Workbooks.Add(visibleWorkbook);
            application.ActiveWorkbook = visibleWorkbook;
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeSchedule(scheduler, @"C:\cases\closed.xlsx", @"C:\cases");
            InvokeExecutePendingPostCloseQueue(scheduler);

            Assert.Equal(0, application.QuitCallCount);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-decision")
                    && ContainsFragment(message, @"workbook=C:\cases\closed.xlsx")
                    && ContainsFragment(message, "targetWorkbookStillOpen=False")
                    && ContainsFragment(message, "decision=scan-visible-workbooks")
                    && ContainsFragment(message, "workbooksCount=1")
                    && ContainsFragment(message, "activeWorkbookPresent=True"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionNotRequired")
                    && ContainsFragment(message, "hasVisibleWorkbook=True")
                    && ContainsFragment(message, "quitAttempted=False")
                    && ContainsFragment(message, "quitCompleted=False")
                    && ContainsFragment(message, "outcomeReason=visibleWorkbookExists")
                    && ContainsFragment(message, "workbooksCount=1")
                    && ContainsFragment(message, "activeWorkbookPresent=True")
                    && ContainsFragment(message, "targetWorkbookStillOpen=unknown"));
            Assert.False(loggerMessages.Exists(message => ContainsFragment(message, "WhiteExcelPreventionSkipped")));
        }

        [Fact]
        public void ScheduleManagedWorkbookClose_ForAccountingClose_WritesShortTtlMarkerAndLogsKind()
        {
            var loggerMessages = new List<string>();
            var application = new Excel.Application();
            string markerPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".marker");
            var markerStore = new ManagedWorkbookCloseMarkerStore(markerPath, () => new DateTime(2026, 5, 12, 0, 0, 0, DateTimeKind.Utc));
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages), markerStore);

            InvokeScheduleManagedWorkbookClose(
                scheduler,
                @"C:\cases\accounting.xlsx",
                @"C:\cases",
                ManagedWorkbookCloseMarkerKind.AccountingClose);

            ManagedWorkbookCloseMarkerReadResult markerResult = markerStore.Consume();
            Assert.True(markerResult.IsValid);
            Assert.Equal(ManagedWorkbookCloseMarkerKind.AccountingClose, markerResult.Marker.Kind);
            Assert.Equal(ManagedWorkbookCloseMarkerStore.DefaultTimeToLiveSeconds, markerResult.Marker.TimeToLiveSeconds);
            Assert.Equal(@"C:\cases\accounting.xlsx", markerResult.Marker.WorkbookKey);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=managed-close-marker-written")
                    && ContainsFragment(message, "managedCloseKind=AccountingClose")
                    && ContainsFragment(message, "ttlSeconds=15"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionQueued")
                    && ContainsFragment(message, "managedCloseKind=AccountingClose"));
        }

        [Fact]
        public void ManagedWorkbookCloseMarkerStore_Consume_WhenExpired_RemovesMarker()
        {
            string markerPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".marker");
            DateTime now = new DateTime(2026, 5, 12, 0, 0, 0, DateTimeKind.Utc);
            var markerStore = new ManagedWorkbookCloseMarkerStore(markerPath, () => now);
            markerStore.Write(ManagedWorkbookCloseMarkerKind.CaseClose, @"C:\cases\case.xlsx");

            now = now.AddSeconds(ManagedWorkbookCloseMarkerStore.DefaultTimeToLiveSeconds + 1);
            ManagedWorkbookCloseMarkerReadResult result = markerStore.Consume();

            Assert.Equal(ManagedWorkbookCloseMarkerReadStatus.Expired, result.Status);
            Assert.False(File.Exists(markerPath));
        }

        private static object CreateScheduler(Excel.Application application, Logger logger)
        {
            return CreateScheduler(application, logger, null);
        }

        private static object CreateScheduler(Excel.Application application, Logger logger, ManagedWorkbookCloseMarkerStore markerStore)
        {
            Assembly addInAssembly = typeof(PostCloseFollowUpScheduler).Assembly;
            Type pathCompatibilityServiceType = addInAssembly.GetType("CaseInfoSystem.ExcelAddIn.Infrastructure.PathCompatibilityService", throwOnError: true);
            Type excelInteropServiceType = addInAssembly.GetType("CaseInfoSystem.ExcelAddIn.Infrastructure.ExcelInteropService", throwOnError: true);

            object pathCompatibilityService = Activator.CreateInstance(pathCompatibilityServiceType, new object[] { null });
            object excelInteropService = Activator.CreateInstance(excelInteropServiceType, new[] { application, logger, pathCompatibilityService });

            return Activator.CreateInstance(
                typeof(PostCloseFollowUpScheduler),
                BindingFlags.Instance | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { application, excelInteropService, logger, markerStore },
                culture: null);
        }

        private static void InvokeQuitExcelIfNoVisibleWorkbook(object scheduler)
        {
            MethodInfo method = typeof(PostCloseFollowUpScheduler).GetMethod(
                "QuitExcelIfNoVisibleWorkbook",
                BindingFlags.Instance | BindingFlags.NonPublic);

            try
            {
                method.Invoke(scheduler, Array.Empty<object>());
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
        }

        private static void InvokeSchedule(object scheduler, string workbookKey, string folderPath)
        {
            MethodInfo method = typeof(PostCloseFollowUpScheduler).GetMethod(
                "Schedule",
                BindingFlags.Instance | BindingFlags.NonPublic);

            method.Invoke(scheduler, new object[] { workbookKey, folderPath });
        }

        private static void InvokeScheduleManagedWorkbookClose(
            object scheduler,
            string workbookKey,
            string folderPath,
            ManagedWorkbookCloseMarkerKind closeKind)
        {
            MethodInfo method = typeof(PostCloseFollowUpScheduler).GetMethod(
                "ScheduleManagedWorkbookClose",
                BindingFlags.Instance | BindingFlags.NonPublic);

            method.Invoke(scheduler, new object[] { workbookKey, folderPath, closeKind });
        }

        private static void InvokeExecutePendingPostCloseQueue(object scheduler)
        {
            MethodInfo method = typeof(PostCloseFollowUpScheduler).GetMethod(
                "ExecutePendingPostCloseQueue",
                BindingFlags.Instance | BindingFlags.NonPublic);

            try
            {
                method.Invoke(scheduler, Array.Empty<object>());
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
        }

        private static bool ContainsFragment(string message, string fragment)
        {
            return message != null
                && fragment != null
                && message.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void AddWorkbookWindow(Excel.Workbook workbook, bool visible)
        {
            workbook.Windows.Add(new Excel.Window { Visible = visible });
        }
    }
}
