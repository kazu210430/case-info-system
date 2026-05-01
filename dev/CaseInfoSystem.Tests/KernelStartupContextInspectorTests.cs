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
        public void Inspect_CapturesStartupWorkbookAndActiveWorkbook()
        {
            var application = new Excel.Application();
            var activeWorkbook = new Excel.Workbook
            {
                FullName = @"C:\root\案件情報System_Kernel.xlsm",
                Name = "案件情報System_Kernel.xlsm"
            };
            application.Workbooks.Add(activeWorkbook);
            application.ActiveWorkbook = activeWorkbook;
            var startupWorkbook = new Excel.Workbook
            {
                FullName = @"C:\root\案件情報System_Base.xlsx",
                Name = "案件情報System_Base.xlsx"
            };

            KernelStartupContextInspector inspector = CreateInspector(
                application,
                getOpenKernelWorkbook: () => null,
                hasVisibleNonKernelWorkbook: () => false);

            KernelStartupContext context = inspector.Inspect(startupWorkbook);

            Assert.Same(startupWorkbook, context.StartupWorkbook);
            Assert.Same(activeWorkbook, context.ActiveWorkbook);
            Assert.False(context.ActiveWorkbookAccessFailed);
        }

        [Fact]
        public void PopulateOpenKernelWorkbookState_UsesLocatorResult()
        {
            var application = new Excel.Application();
            var kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\root\案件情報System_Kernel.xlsm",
                Name = "案件情報System_Kernel.xlsm"
            };
            KernelStartupContextInspector inspector = CreateInspector(
                application,
                getOpenKernelWorkbook: () => kernelWorkbook,
                hasVisibleNonKernelWorkbook: () => false);
            KernelStartupContext context = inspector.Inspect(startupWorkbook: null);

            inspector.PopulateOpenKernelWorkbookState(context);

            Assert.True(context.HasOpenKernelWorkbook);
        }

        [Fact]
        public void PopulateVisibleNonKernelWorkbookState_UsesVisibilityResolverResult()
        {
            var application = new Excel.Application();
            KernelStartupContextInspector inspector = CreateInspector(
                application,
                getOpenKernelWorkbook: () => null,
                hasVisibleNonKernelWorkbook: () => true);
            KernelStartupContext context = inspector.Inspect(startupWorkbook: null);

            inspector.PopulateVisibleNonKernelWorkbookState(context);

            Assert.True(context.HasVisibleNonKernelWorkbook);
        }

        private static KernelStartupContextInspector CreateInspector(
            Excel.Application application,
            System.Func<Excel.Workbook> getOpenKernelWorkbook,
            System.Func<bool> hasVisibleNonKernelWorkbook)
        {
            Logger logger = OrchestrationTestSupport.CreateLogger(new List<string>());
            var locator = new KernelOpenWorkbookLocator(
                application,
                excelInteropService: null,
                pathCompatibilityService: new PathCompatibilityService(),
                logger,
                getOpenKernelWorkbookOverride: getOpenKernelWorkbook);

            return new KernelStartupContextInspector(
                application,
                locator,
                hasVisibleNonKernelWorkbook);
        }
    }
}
