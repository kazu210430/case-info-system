using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookStartupDisplayPolicyTests
    {
        [Fact]
        public void ShouldShowHomeOnStartup_ReturnsFalse_WhenKernelStartupContextIsMissing()
        {
            bool result = KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext: false,
                hasKernelWorkbookContext: true,
                isStartupWorkbookKernel: true,
                hasVisibleNonKernelWorkbook: false);

            Assert.False(result);
        }

        [Fact]
        public void ShouldShowHomeOnStartup_ReturnsFalse_WhenKernelWorkbookContextIsMissing()
        {
            bool result = KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext: true,
                hasKernelWorkbookContext: false,
                isStartupWorkbookKernel: true,
                hasVisibleNonKernelWorkbook: false);

            Assert.False(result);
        }

        [Fact]
        public void ShouldShowHomeOnStartup_ReturnsTrue_WhenStartupWorkbookIsKernel()
        {
            bool result = KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext: true,
                hasKernelWorkbookContext: true,
                isStartupWorkbookKernel: true,
                hasVisibleNonKernelWorkbook: true);

            Assert.True(result);
        }

        [Theory]
        [InlineData(false, true)]
        [InlineData(true, false)]
        public void ShouldShowHomeOnStartup_UsesVisibleNonKernelWorkbookAsFallback(
            bool hasVisibleNonKernelWorkbook,
            bool expected)
        {
            bool result = KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext: true,
                hasKernelWorkbookContext: true,
                isStartupWorkbookKernel: false,
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook);

            Assert.Equal(expected, result);
        }
    }
}
