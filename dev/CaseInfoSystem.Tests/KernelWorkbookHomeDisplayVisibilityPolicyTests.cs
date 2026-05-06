using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookHomeDisplayVisibilityPolicyTests
    {
        [Fact]
        public void DecideAction_ReturnsMinimizeKernelWindows_WhenVisibleNonKernelWorkbookExists()
        {
            KernelWorkbookHomeDisplayVisibilityAction result = KernelWorkbookHomeDisplayVisibilityPolicy.DecideAction(
                hasVisibleNonKernelWorkbook: true,
                isActiveWorkbookKernel: true,
                visibleWorkbookCount: 1);

            Assert.Equal(KernelWorkbookHomeDisplayVisibilityAction.MinimizeKernelWindows, result);
        }

        [Fact]
        public void DecideAction_ReturnsConcealKernelWindowsAndHideExcelMainWindow_WhenActiveKernelWorkbookIsVisible()
        {
            KernelWorkbookHomeDisplayVisibilityAction result = KernelWorkbookHomeDisplayVisibilityPolicy.DecideAction(
                hasVisibleNonKernelWorkbook: false,
                isActiveWorkbookKernel: true,
                visibleWorkbookCount: 1);

            Assert.Equal(KernelWorkbookHomeDisplayVisibilityAction.ConcealKernelWindowsAndHideExcelMainWindow, result);
        }

        [Theory]
        [InlineData(false, false, 1)]
        [InlineData(false, true, 0)]
        [InlineData(false, true, -1)]
        public void DecideAction_ReturnsHideExcelMainWindowOnly_WhenConcealPreconditionsAreNotMet(
            bool hasVisibleNonKernelWorkbook,
            bool isActiveWorkbookKernel,
            int visibleWorkbookCount)
        {
            KernelWorkbookHomeDisplayVisibilityAction result = KernelWorkbookHomeDisplayVisibilityPolicy.DecideAction(
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook,
                isActiveWorkbookKernel: isActiveWorkbookKernel,
                visibleWorkbookCount: visibleWorkbookCount);

            Assert.Equal(KernelWorkbookHomeDisplayVisibilityAction.HideExcelMainWindowOnly, result);
        }
    }
}
