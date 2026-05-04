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
        public void HasAnyOpenKernelWorkbook_ReturnsTrueWithoutCallingGetOpenOverride_WhenKernelWorkbookExists()
        {
            var application = new Excel.Application();
            AddWorkbook(application, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
            int getOpenCalls = 0;
            KernelOpenWorkbookLocator locator = CreateLocator(
                application,
                () =>
                {
                    getOpenCalls++;
                    return null;
                });

            bool result = locator.HasAnyOpenKernelWorkbook();

            Assert.True(result);
            Assert.Equal(0, getOpenCalls);
        }

        [Fact]
        public void HasAnyOpenKernelWorkbook_ReturnsFalseWithoutCallingGetOpenOverride_WhenKernelWorkbookIsAbsent()
        {
            var application = new Excel.Application();
            AddWorkbook(application, "ExistingCase.xlsx");
            int getOpenCalls = 0;
            KernelOpenWorkbookLocator locator = CreateLocator(
                application,
                () =>
                {
                    getOpenCalls++;
                    return new Excel.Workbook
                    {
                        Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm")
                    };
                });

            bool result = locator.HasAnyOpenKernelWorkbook();

            Assert.False(result);
            Assert.Equal(0, getOpenCalls);
        }

        private static KernelOpenWorkbookLocator CreateLocator(Excel.Application application, System.Func<Excel.Workbook> getOpenKernelWorkbookOverride)
        {
            var loggerMessages = new List<string>();
            Logger logger = OrchestrationTestSupport.CreateLogger(loggerMessages);
            return new KernelOpenWorkbookLocator(
                application,
                null,
                new PathCompatibilityService(),
                logger,
                getOpenKernelWorkbookOverride: getOpenKernelWorkbookOverride);
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
