using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class KernelOpenWorkbookLocatorTests
    {
        [Fact]
        public void HasAnyOpenKernelWorkbook_ReturnsTrue_WhenKernelWorkbookExists()
        {
            var application = new Excel.Application();
            AddWorkbook(application, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
            KernelOpenWorkbookLocator locator = CreateLocator(application);

            bool result = locator.HasAnyOpenKernelWorkbook();

            Assert.True(result);
        }

        [Fact]
        public void HasAnyOpenKernelWorkbook_ReturnsFalse_WhenKernelWorkbookIsAbsent()
        {
            var application = new Excel.Application();
            AddWorkbook(application, "ExistingCase.xlsx");
            KernelOpenWorkbookLocator locator = CreateLocator(application);

            bool result = locator.HasAnyOpenKernelWorkbook();

            Assert.False(result);
        }

        private static KernelOpenWorkbookLocator CreateLocator(Excel.Application application)
        {
            var loggerMessages = new List<string>();
            Logger logger = OrchestrationTestSupport.CreateLogger(loggerMessages);
            return new KernelOpenWorkbookLocator(
                application,
                null,
                new PathCompatibilityService(),
                logger);
        }

        private static Excel.Workbook AddWorkbook(Excel.Application application, string workbookName)
        {
            var workbook = new Excel.Workbook
            {
                Name = workbookName,
                FullName = "C:\\Work\\" + workbookName
            };

            application.Workbooks.Add(workbook);
            return workbook;
        }
    }
}
