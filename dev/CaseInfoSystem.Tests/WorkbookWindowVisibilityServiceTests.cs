using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    [CollectionDefinition("ExcelApplicationCreatedApplications", DisableParallelization = true)]
    public sealed class ExcelApplicationCreatedApplicationsCollection
    {
    }

    [Collection("ExcelApplicationCreatedApplications")]
    public class WorkbookWindowVisibilityServiceTests : System.IDisposable
    {
        public WorkbookWindowVisibilityServiceTests()
        {
            Excel.Application.ResetCreatedApplications();
        }

        public void Dispose()
        {
            Excel.Application.ResetCreatedApplications();
        }

        [Fact]
        public void EnsureVisible_WhenWorkbookWindowIsAlreadyVisible_ReturnsAlreadyVisible()
        {
            WorkbookWindowVisibilityService service = CreateService(out Excel.Workbook workbook, out Excel.Window window);

            window.Visible = true;

            WorkbookWindowVisibilityEnsureResult result = service.EnsureVisible(workbook, "test");

            Assert.Equal(WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible, result.Outcome);
            Assert.True(window.Visible);
            Assert.Equal("101", result.WindowHwnd);
        }

        [Fact]
        public void EnsureVisible_WhenOnlyHiddenWorkbookWindowExists_MakesWindowVisible()
        {
            WorkbookWindowVisibilityService service = CreateService(out Excel.Workbook workbook, out Excel.Window window);

            window.Visible = false;

            WorkbookWindowVisibilityEnsureResult result = service.EnsureVisible(workbook, "test");

            Assert.Equal(WorkbookWindowVisibilityEnsureOutcome.MadeVisible, result.Outcome);
            Assert.True(window.Visible);
            Assert.True(result.VisibleAfterSet);
            Assert.Equal("101", result.WindowHwnd);
        }

        [Fact]
        public void EnsureVisible_WhenWorkbookWindowCannotBeResolved_ReturnsWindowUnresolved()
        {
            WorkbookWindowVisibilityService service = CreateService(out Excel.Workbook workbook, out _);

            workbook.Windows.Clear();

            WorkbookWindowVisibilityEnsureResult result = service.EnsureVisible(workbook, "test");

            Assert.Equal(WorkbookWindowVisibilityEnsureOutcome.WindowUnresolved, result.Outcome);
            Assert.Equal(string.Empty, result.WindowHwnd);
        }

        private static WorkbookWindowVisibilityService CreateService(out Excel.Workbook workbook, out Excel.Window window)
        {
            var application = new Excel.Application();
            workbook = new Excel.Workbook
            {
                Application = application,
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
            };
            application.Workbooks.Add(workbook);
            window = workbook.Windows[1];
            window.Hwnd = 101;

            var logger = OrchestrationTestSupport.CreateLogger(new List<string>());
            var excelInteropService = new ExcelInteropService(application, logger, new PathCompatibilityService());
            return new WorkbookWindowVisibilityService(excelInteropService, logger);
        }
    }
}
