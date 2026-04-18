using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookPromotionPolicyTests
    {
        [Fact]
        public void ShouldPromoteKernelWorkbookOnHomeRelease_ReturnsTrue_DuringCaseCreationFlow()
        {
            bool result = KernelWorkbookPromotionPolicy.ShouldPromoteKernelWorkbookOnHomeRelease(
                isKernelCaseCreationFlowActive: true,
                hasActiveWorkbook: true,
                isActiveWorkbookKernel: false,
                hasVisibleNonKernelWorkbook: true);

            Assert.True(result);
        }

        [Theory]
        [InlineData(false, true)]
        [InlineData(true, false)]
        public void ShouldPromoteKernelWorkbookOnHomeRelease_UsesVisibleNonKernelWorkbook_WhenActiveWorkbookIsMissing(
            bool hasVisibleNonKernelWorkbook,
            bool expected)
        {
            bool result = KernelWorkbookPromotionPolicy.ShouldPromoteKernelWorkbookOnHomeRelease(
                isKernelCaseCreationFlowActive: false,
                hasActiveWorkbook: false,
                isActiveWorkbookKernel: false,
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook);

            Assert.Equal(expected, result);
        }

        [Fact]
        public void ShouldPromoteKernelWorkbookOnHomeRelease_ReturnsTrue_WhenActiveWorkbookIsKernel()
        {
            bool result = KernelWorkbookPromotionPolicy.ShouldPromoteKernelWorkbookOnHomeRelease(
                isKernelCaseCreationFlowActive: false,
                hasActiveWorkbook: true,
                isActiveWorkbookKernel: true,
                hasVisibleNonKernelWorkbook: false);

            Assert.True(result);
        }

        [Fact]
        public void ShouldPromoteKernelWorkbookOnHomeRelease_ReturnsFalse_WhenActiveWorkbookIsNonKernel()
        {
            bool result = KernelWorkbookPromotionPolicy.ShouldPromoteKernelWorkbookOnHomeRelease(
                isKernelCaseCreationFlowActive: false,
                hasActiveWorkbook: true,
                isActiveWorkbookKernel: false,
                hasVisibleNonKernelWorkbook: false);

            Assert.False(result);
        }
    }
}
