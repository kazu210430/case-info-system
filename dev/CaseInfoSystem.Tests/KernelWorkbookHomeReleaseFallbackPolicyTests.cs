using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookHomeReleaseFallbackPolicyTests
    {
        [Theory]
        [InlineData(true, false, 0)]
        [InlineData(true, true, 0)]
        [InlineData(false, false, 1)]
        [InlineData(false, true, 2)]
        public void DecideHomeReleaseAction_ReturnsExpectedAction(
            bool shouldAvoidGlobalExcelWindowRestore,
            bool shouldPromoteKernelWorkbook,
            int expected)
        {
            KernelWorkbookHomeReleaseAction result = KernelWorkbookHomeReleaseFallbackPolicy.DecideHomeReleaseAction(
                shouldAvoidGlobalExcelWindowRestore: shouldAvoidGlobalExcelWindowRestore,
                shouldPromoteKernelWorkbook: shouldPromoteKernelWorkbook);

            Assert.Equal(expected, (int)result);
        }
    }
}
