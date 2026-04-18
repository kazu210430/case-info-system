using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookWindowRestorePolicyTests
    {
        [Theory]
        [InlineData(false, false, false)]
        [InlineData(false, true, true)]
        [InlineData(true, false, false)]
        [InlineData(true, true, false)]
        public void ShouldAvoidGlobalExcelWindowRestore_UsesCaseCreationFlowAndVisibleNonKernelWorkbook(
            bool isKernelCaseCreationFlowActive,
            bool hasVisibleNonKernelWorkbook,
            bool expected)
        {
            bool result = KernelWorkbookWindowRestorePolicy.ShouldAvoidGlobalExcelWindowRestore(
                isKernelCaseCreationFlowActive: isKernelCaseCreationFlowActive,
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook);

            Assert.Equal(expected, result);
        }
    }
}
