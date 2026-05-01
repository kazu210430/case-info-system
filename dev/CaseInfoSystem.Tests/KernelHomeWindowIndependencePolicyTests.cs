using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelHomeWindowIndependencePolicyTests
    {
        [Fact]
        public void ShouldCloseKernelHome_ReturnsFalse_WhenKernelCaseCreationFlowIsInactive()
        {
            bool result = KernelHomeExternalClosePolicy.ShouldCloseKernelHome(
                isKernelCaseCreationFlowActive: false);

            Assert.False(result);
        }

        [Fact]
        public void ShouldCloseKernelHome_ReturnsTrue_WhenKernelCaseCreationFlowIsActive()
        {
            bool result = KernelHomeExternalClosePolicy.ShouldCloseKernelHome(
                isKernelCaseCreationFlowActive: true);

            Assert.True(result);
        }

        [Theory]
        [InlineData("WorkbookOpen", true, true)]
        [InlineData("WorkbookActivate", true, false)]
        [InlineData("WorkbookOpen", false, false)]
        [InlineData("SheetActivate", true, false)]
        public void ShouldAutoShow_RestrictsKernelHomeAutoDisplayToWorkbookOpen(
            string eventName,
            bool startupPolicyAllowsDisplay,
            bool expected)
        {
            bool result = KernelHomeAutoDisplayEventPolicy.ShouldAutoShow(
                eventName,
                startupPolicyAllowsDisplay);

            Assert.Equal(expected, result);
        }
    }
}
