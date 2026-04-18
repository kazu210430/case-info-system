using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class TransientPaneSuppressionServiceTests
    {
        [Fact]
        public void Suppression_Add_Remove_WorksCorrectly()
        {
            var service = CreateService();
            const string workbookPath = @"C:\Cases\1001\Case.xlsx";

            service.SuppressPath(workbookPath, "test");

            Assert.True(service.IsSuppressedPath(workbookPath));

            service.ReleasePath(workbookPath, "test");

            Assert.False(service.IsSuppressedPath(workbookPath));
        }

        private static TransientPaneSuppressionService CreateService()
        {
            return new TransientPaneSuppressionService(
                new FakeExcelInteropService(),
                new PathCompatibilityService(),
                new Logger(_ => { }));
        }
    }
}
