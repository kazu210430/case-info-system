using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class KernelStartupContextInspectorTests
    {
        [Fact]
        public void InspectForStartupDisplay_CollectsOpenKernelWorkbookFact_WhenStartupWorkbookIsKernel()
        {
            var application = new Excel.Application();
            Excel.Workbook startupWorkbook = AddWorkbook(application, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
            KernelStartupContextInspector inspector = CreateInspector(application);

            KernelStartupContext context = inspector.InspectForStartupDisplay(startupWorkbook);

            Assert.Same(startupWorkbook, context.StartupWorkbook);
            Assert.True(context.StartupWorkbookIsKernel);
            Assert.False(context.StartupContextActiveWorkbookIsKernel);
            Assert.False(context.StartupContextActiveWorkbookReadFailed);
            Assert.False(context.KernelContextActiveWorkbookIsKernel);
            Assert.False(context.KernelContextActiveWorkbookReadFailed);
            Assert.True(context.HasOpenKernelWorkbook);
            Assert.False(context.HasVisibleNonKernelWorkbook);
        }

        [Fact]
        public void InspectForStartupDisplay_CollectsVisibleNonKernelWorkbookFact_WhenActiveKernelProvidesContext()
        {
            var application = new Excel.Application();
            Excel.Workbook startupWorkbook = AddWorkbook(application, "ExistingCase.xlsx");
            Excel.Workbook kernelWorkbook = AddWorkbook(application, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
            application.ActiveWorkbook = kernelWorkbook;
            KernelStartupContextInspector inspector = CreateInspector(application);

            KernelStartupContext context = inspector.InspectForStartupDisplay(startupWorkbook);

            Assert.False(context.StartupWorkbookIsKernel);
            Assert.True(context.StartupContextActiveWorkbookIsKernel);
            Assert.False(context.StartupContextActiveWorkbookReadFailed);
            Assert.True(context.KernelContextActiveWorkbookIsKernel);
            Assert.False(context.KernelContextActiveWorkbookReadFailed);
            Assert.False(context.HasOpenKernelWorkbook);
            Assert.True(context.HasVisibleNonKernelWorkbook);
        }

        [Fact]
        public void InspectForStartupDescription_CollectsDescriptionFacts_WithoutChangingStartupPolicyInputs()
        {
            var application = new Excel.Application();
            Excel.Workbook kernelWorkbook = AddWorkbook(application, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
            AddWorkbook(application, "ExistingCase.xlsx");
            application.ActiveWorkbook = kernelWorkbook;
            KernelStartupContextInspector inspector = CreateInspector(application);

            KernelStartupContext context = inspector.InspectForStartupDescription();

            Assert.Equal(kernelWorkbook.Name, context.DescribeActiveWorkbookName);
            Assert.True(context.DescribeActiveWorkbookIsKernel);
            Assert.False(context.DescribeActiveWorkbookReadFailed);
            Assert.True(context.HasOpenKernelWorkbook);
            Assert.True(context.HasVisibleNonKernelWorkbook);
        }

        [Fact]
        public void InspectForStartupDescription_ReportsNoOpenKernelWorkbook_WhenKernelIsAbsent()
        {
            var application = new Excel.Application();
            Excel.Workbook activeWorkbook = AddWorkbook(application, "ExistingCase.xlsx");
            application.ActiveWorkbook = activeWorkbook;
            KernelStartupContextInspector inspector = CreateInspector(application);

            KernelStartupContext context = inspector.InspectForStartupDescription();

            Assert.Equal(activeWorkbook.Name, context.DescribeActiveWorkbookName);
            Assert.False(context.DescribeActiveWorkbookIsKernel);
            Assert.False(context.DescribeActiveWorkbookReadFailed);
            Assert.False(context.HasOpenKernelWorkbook);
            Assert.True(context.HasVisibleNonKernelWorkbook);
        }

        private static KernelStartupContextInspector CreateInspector(Excel.Application application)
        {
            var loggerMessages = new List<string>();
            Logger logger = OrchestrationTestSupport.CreateLogger(loggerMessages);
            var locator = new KernelOpenWorkbookLocator(application, null, new PathCompatibilityService(), logger);
            return new KernelStartupContextInspector(application, null, locator);
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
