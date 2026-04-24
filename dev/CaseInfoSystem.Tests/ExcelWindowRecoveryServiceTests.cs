using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class ExcelWindowRecoveryServiceTests
    {
        [Fact]
        public void TryRecoverWorkbookWindow_WhenWorkbookWindowIsMissing_RecreatesWorkbookWindow()
        {
            var logMessages = new List<string>();
            var application = new Excel.Application
            {
                Visible = false,
                ScreenUpdating = true
            };
            var workbook = new Excel.Workbook
            {
                Application = application,
                FullName = @"C:\root\kernel.xlsm",
                Name = "kernel.xlsm"
            };
            application.Workbooks.Add(workbook);

            var logger = OrchestrationTestSupport.CreateLogger(logMessages);
            var excelInteropService = new ExcelInteropService(
                application,
                logger,
                new PathCompatibilityService());
            var service = new ExcelWindowRecoveryService(
                application,
                excelInteropService,
                logger);

            bool recovered = service.TryRecoverWorkbookWindow(workbook, "test", bringToFront: false);

            Assert.True(recovered);
            Assert.True(application.Visible);
            Assert.Equal(1, workbook.Windows.Count);
            Assert.True(workbook.Windows[1].Visible);
            Assert.True(workbook.Windows[1].Activated);
        }
    }
}
