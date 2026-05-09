using System;
using System.Collections.Generic;
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

            hiddenWorkbook.Windows[1].Visible = false;
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

            hiddenWorkbook.Windows[1].Visible = false;
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

            visibleWorkbook.Windows[1].Visible = true;
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
                    && ContainsFragment(message, "folderPathPresent=True")
                    && ContainsFragment(message, "targetWorkbookStillOpen=unknown"));
        }

        [Fact]
        public void ExecutePendingPostCloseQueue_WhenTargetWorkbookStillOpen_LogsDecisionAndDoesNotQuit()
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
                    && ContainsFragment(message, "attemptsRemaining=20"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-decision")
                    && ContainsFragment(message, "targetWorkbookStillOpen=True")
                    && ContainsFragment(message, "decision=skip-still-open"));
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "WhiteExcelPreventionNotRequired")
                    && ContainsFragment(message, "hasVisibleWorkbook=unknown")
                    && ContainsFragment(message, "quitAttempted=False")
                    && ContainsFragment(message, "quitCompleted=False")
                    && ContainsFragment(message, "outcomeReason=targetWorkbookStillOpen")
                    && ContainsFragment(message, "targetWorkbookStillOpen=True"));
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

            visibleWorkbook.Windows[1].Visible = true;
            application.Workbooks.Add(visibleWorkbook);
            object scheduler = CreateScheduler(application, OrchestrationTestSupport.CreateLogger(loggerMessages));

            InvokeSchedule(scheduler, @"C:\cases\closed.xlsx", @"C:\cases");
            InvokeExecutePendingPostCloseQueue(scheduler);

            Assert.Equal(0, application.QuitCallCount);
            Assert.Contains(
                loggerMessages,
                message => ContainsFragment(message, "action=post-close-follow-up-decision")
                    && ContainsFragment(message, @"workbook=C:\cases\closed.xlsx")
                    && ContainsFragment(message, "targetWorkbookStillOpen=False")
                    && ContainsFragment(message, "decision=scan-visible-workbooks"));
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

        private static object CreateScheduler(Excel.Application application, Logger logger)
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
                args: new[] { application, excelInteropService, logger },
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
    }
}
