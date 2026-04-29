using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelHomeDisplayAvailabilityPolicyTests
    {
        [Fact]
        public void Decide_ReturnsShow_WhenKernelReachedDisplayIsReadyAndContextAllowsIt()
        {
            KernelHomeDisplayAction action = KernelHomeDisplayAvailabilityPolicy.Decide(
                hasKernelWorkbookReached: true,
                isDisplayReady: true,
                hasVisibleKernelHome: false,
                isSuppressed: false,
                isDisplayContextAllowed: true,
                shouldReloadVisibleKernelHome: false);

            Assert.Equal(KernelHomeDisplayAction.Show, action);
        }

        [Theory]
        [InlineData(false, true, false)]
        [InlineData(true, false, false)]
        [InlineData(true, true, true)]
        public void Decide_ReturnsNone_WhenShowPreconditionsAreNotMet(
            bool hasKernelWorkbookReached,
            bool isDisplayReady,
            bool isSuppressed)
        {
            KernelHomeDisplayAction action = KernelHomeDisplayAvailabilityPolicy.Decide(
                hasKernelWorkbookReached,
                isDisplayReady,
                hasVisibleKernelHome: false,
                isSuppressed,
                isDisplayContextAllowed: true,
                shouldReloadVisibleKernelHome: false);

            Assert.Equal(KernelHomeDisplayAction.None, action);
        }

        [Fact]
        public void Decide_ReturnsReloadVisible_WhenVisibleHomeShouldBeReloaded()
        {
            KernelHomeDisplayAction action = KernelHomeDisplayAvailabilityPolicy.Decide(
                hasKernelWorkbookReached: true,
                isDisplayReady: false,
                hasVisibleKernelHome: true,
                isSuppressed: false,
                isDisplayContextAllowed: false,
                shouldReloadVisibleKernelHome: true);

            Assert.Equal(KernelHomeDisplayAction.ReloadVisible, action);
        }
    }
}
