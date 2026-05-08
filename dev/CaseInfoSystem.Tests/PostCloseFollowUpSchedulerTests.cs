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
            Assert.Contains(loggerMessages, message => message.IndexOf("WhiteExcelPreventionCompleted", StringComparison.OrdinalIgnoreCase) >= 0);
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
            Assert.Contains(loggerMessages, message => message.IndexOf("WhiteExcelPreventionFailed", StringComparison.OrdinalIgnoreCase) >= 0);
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
            Assert.Contains(loggerMessages, message => message.IndexOf("WhiteExcelPreventionNotRequired", StringComparison.OrdinalIgnoreCase) >= 0);
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
    }
}
