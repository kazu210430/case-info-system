using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneHostReusePolicyTests
    {
        [Theory]
        [InlineData("WorkbookActivate")]
        [InlineData("windowactivate")]
        [InlineData("KernelHomeForm.FormClosed")]
        public void ShouldReuseCaseHostWithoutRender_ReturnsTrue_ForRenderedCaseHostOnSameWorkbookAndSupportedReason(string reason)
        {
            bool result = TaskPaneHostReusePolicy.ShouldReuseCaseHostWithoutRender(
                WorkbookRole.Case,
                isDocumentButtonsHost: true,
                isAlreadyRendered: true,
                isSameWorkbook: true,
                reason: reason);

            Assert.True(result);
        }

        [Fact]
        public void ShouldReuseCaseHostWithoutRender_ReturnsFalse_ForNonCaseRole()
        {
            bool result = TaskPaneHostReusePolicy.ShouldReuseCaseHostWithoutRender(
                WorkbookRole.Kernel,
                isDocumentButtonsHost: true,
                isAlreadyRendered: true,
                isSameWorkbook: true,
                reason: "WorkbookActivate");

            Assert.False(result);
        }

        [Theory]
        [InlineData(false, true, "WorkbookActivate")]
        [InlineData(true, false, "WorkbookActivate")]
        [InlineData(true, true, "WorkbookOpen")]
        [InlineData(true, true, null)]
        public void ShouldReuseCaseHostWithoutRender_ReturnsFalse_WhenRequiredStateIsMissing(
            bool isAlreadyRendered,
            bool isSameWorkbook,
            string reason)
        {
            bool result = TaskPaneHostReusePolicy.ShouldReuseCaseHostWithoutRender(
                WorkbookRole.Case,
                isDocumentButtonsHost: true,
                isAlreadyRendered: isAlreadyRendered,
                isSameWorkbook: isSameWorkbook,
                reason: reason);

            Assert.False(result);
        }

        [Fact]
        public void ShouldReuseCaseHostWithoutRender_ReturnsFalse_ForNonDocumentButtonsHost()
        {
            bool result = TaskPaneHostReusePolicy.ShouldReuseCaseHostWithoutRender(
                WorkbookRole.Case,
                isDocumentButtonsHost: false,
                isAlreadyRendered: true,
                isSameWorkbook: true,
                reason: "WorkbookActivate");

            Assert.False(result);
        }
    }
}
