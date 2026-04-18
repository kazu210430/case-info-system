using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class NavigationServiceTests
    {
        [Fact]
        public void CreateContext_SetsExpectedProperties()
        {
            var service = new NavigationService(
                new FakeExcelInteropService(),
                new FakeWorkbookRoleResolver(),
                CreateLogger());

            var context = service.CreateContextFromSeed(
                new WorkbookContextSeed(
                    workbook: null,
                    window: null,
                    role: WorkbookRole.Accounting,
                    systemRoot: @"C:\Cases\1001",
                    workbookFullName: @"C:\Cases\1001\Accounting.xlsx",
                    activeSheetCodeName: "shHOME"));

            Assert.NotNull(context);
            Assert.Equal(WorkbookRole.Accounting, context.Role);
            Assert.Equal(@"C:\Cases\1001", context.SystemRoot);
            Assert.Equal(@"C:\Cases\1001\Accounting.xlsx", context.WorkbookFullName);
            Assert.Equal("shHOME", context.ActiveSheetCodeName);
            Assert.Null(context.Workbook);
            Assert.Null(context.Window);
        }

        private static Logger CreateLogger()
        {
            return new Logger(_ => { });
        }
    }
}
