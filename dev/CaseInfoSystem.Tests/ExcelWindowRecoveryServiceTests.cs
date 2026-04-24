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
        public void TryRecoverWorkbookWindow_WhenWorkbookWindowIsMissing_UsesActivationToMaterializeWorkbookWindow()
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

        [Fact]
        public void NormalizeWorkbookWindows_WhenSiblingWindowsExist_KeepsActiveWindowAndClosesOthers()
        {
            var logMessages = new List<string>();
            var application = new Excel.Application
            {
                Visible = true,
                ScreenUpdating = true
            };
            var workbook = new Excel.Workbook
            {
                Application = application,
                FullName = @"C:\root\kernel.xlsm",
                Name = "kernel.xlsm"
            };
            application.Workbooks.Add(workbook);

            Excel.Window firstWindow = workbook.NewWindow();
            firstWindow.Hwnd = 101;
            firstWindow.Visible = true;

            Excel.Window activeWindow = workbook.NewWindow();
            activeWindow.Hwnd = 202;
            activeWindow.Visible = true;
            application.ActiveWorkbook = workbook;
            application.ActiveWindow = activeWindow;

            var logger = OrchestrationTestSupport.CreateLogger(logMessages);
            var excelInteropService = new ExcelInteropService(
                application,
                logger,
                new PathCompatibilityService());
            var service = new ExcelWindowRecoveryService(
                application,
                excelInteropService,
                logger);

            bool normalized = service.NormalizeWorkbookWindows(workbook, "normalize-test", ensurePrimaryVisible: true, activatePrimary: true, bringToFront: false);

            Assert.True(normalized);
            Assert.Equal(1, workbook.Windows.Count);
            Assert.Same(activeWindow, workbook.Windows[1]);
            Assert.True(activeWindow.Visible);
            Assert.True(activeWindow.Activated);
            Assert.True(firstWindow.Closed);
            Assert.Equal(1, firstWindow.CloseCallCount);
        }

        [Fact]
        public void TryRecoverWorkbookWindowUsingExistingWindows_WhenActiveWindowMatchesWorkbook_ReusesExistingWindow()
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
            var existingWindow = new Excel.Window
            {
                Visible = true
            };
            application.Workbooks.Add(workbook);
            application.ActiveWorkbook = workbook;
            application.ActiveWindow = existingWindow;

            var logger = OrchestrationTestSupport.CreateLogger(logMessages);
            var excelInteropService = new ExcelInteropService(
                application,
                logger,
                new PathCompatibilityService());
            var service = new ExcelWindowRecoveryService(
                application,
                excelInteropService,
                logger);

            bool recovered = service.TryRecoverWorkbookWindowUsingExistingWindows(workbook, "test-existing", bringToFront: false);

            Assert.True(recovered);
            Assert.True(application.Visible);
            Assert.Equal(0, workbook.Windows.Count);
            Assert.True(existingWindow.Activated);
        }

        [Fact]
        public void TryRecoverWorkbookWindowUsingExistingWindows_WhenNoVisibleWindowIsFound_UsesActivationWithoutNewWindow()
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

            bool recovered = service.TryRecoverWorkbookWindowUsingExistingWindows(workbook, "test-no-create", bringToFront: false);

            Assert.True(recovered);
            Assert.True(application.Visible);
            Assert.Equal(1, workbook.Windows.Count);
        }
    }
}
