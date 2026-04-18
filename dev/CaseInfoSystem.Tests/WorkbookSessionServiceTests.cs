using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class WorkbookSessionServiceTests
    {
        [Fact]
        public void ShouldHandleContext_ReturnsTrue_ForCaseWorkbook()
        {
            var service = CreateWorkbookSessionService(out _);
            WorkbookContext context = CreateContext(WorkbookRole.Case, @"C:\Cases\1001\Case.xlsx");

            Assert.True(service.ShouldHandleContext(context));
        }

        [Fact]
        public void ShouldHandleContext_ReturnsFalse_WhenSuppressed()
        {
            var service = CreateWorkbookSessionService(out var suppressionService);
            const string workbookPath = @"C:\Cases\1001\Case.xlsx";
            suppressionService.SuppressPath(workbookPath, "test");
            WorkbookContext context = CreateContext(WorkbookRole.Case, workbookPath);

            Assert.False(service.ShouldHandleContext(context));
        }

        [Fact]
        public void ShouldHandleContext_ReturnsTrue_AfterRelease()
        {
            var service = CreateWorkbookSessionService(out var suppressionService);
            const string workbookPath = @"C:\Cases\1001\Case.xlsx";
            suppressionService.SuppressPath(workbookPath, "test");
            suppressionService.ReleasePath(workbookPath, "test");
            WorkbookContext context = CreateContext(WorkbookRole.Case, workbookPath);

            Assert.True(service.ShouldHandleContext(context));
        }

        [Fact]
        public void ShouldHandleContext_ReturnsFalse_ForNonTarget()
        {
            var service = CreateWorkbookSessionService(out _);
            WorkbookContext context = CreateContext(WorkbookRole.Unknown, @"C:\Cases\1001\Other.xlsx");

            Assert.False(service.ShouldHandleContext(context));
        }

        private static WorkbookSessionService CreateWorkbookSessionService(out TransientPaneSuppressionService suppressionService)
        {
            var navigationService = new NavigationService(
                new FakeExcelInteropService(),
                new FakeWorkbookRoleResolver(),
                new Logger(_ => { }));

            suppressionService = new TransientPaneSuppressionService(
                new FakeExcelInteropService(),
                new PathCompatibilityService(),
                new Logger(_ => { }));

            return new WorkbookSessionService(
                navigationService,
                suppressionService,
                new Logger(_ => { }));
        }

        private static WorkbookContext CreateContext(WorkbookRole role, string workbookFullName)
        {
            var navigationService = new NavigationService(
                new FakeExcelInteropService(),
                new FakeWorkbookRoleResolver(),
                new Logger(_ => { }));

            return navigationService.CreateContextFromSeed(
                new WorkbookContextSeed(
                    workbook: null,
                    window: null,
                    role: role,
                    systemRoot: @"C:\Cases\1001",
                    workbookFullName: workbookFullName,
                    activeSheetCodeName: "shCASE"));
        }
    }
}
