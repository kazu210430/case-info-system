using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ExcelWindowRestoreDecisionPolicyTests
    {
        [Fact]
        public void Decide_WhenWindowIsVisibleAndShownNormal_SkipsRestore()
        {
            ExcelWindowRestoreDecision result = ExcelWindowRestoreDecisionPolicy.Decide(
                windowResolved: true,
                visibleReadSucceeded: true,
                visible: true,
                placementReadSucceeded: true,
                showCmd: ExcelWindowRestoreDecisionPolicy.SwShowNormal);

            Assert.False(result.ShouldRestore);
            Assert.Equal("True", result.RestoreSkipped);
            Assert.Equal("visible-shownormal-not-minimized-not-maximized", result.RestoreSkipReason);
        }

        [Theory]
        [InlineData(ExcelWindowRestoreDecisionPolicy.SwShowMinimized)]
        [InlineData(ExcelWindowRestoreDecisionPolicy.SwMinimize)]
        [InlineData(ExcelWindowRestoreDecisionPolicy.SwShowMinNoActive)]
        [InlineData(ExcelWindowRestoreDecisionPolicy.SwForceMinimize)]
        public void Decide_WhenWindowIsMinimized_Restores(int showCmd)
        {
            ExcelWindowRestoreDecision result = ExcelWindowRestoreDecisionPolicy.Decide(
                windowResolved: true,
                visibleReadSucceeded: true,
                visible: true,
                placementReadSucceeded: true,
                showCmd: showCmd);

            Assert.True(result.ShouldRestore);
            Assert.Equal("False", result.RestoreSkipped);
            Assert.Equal("restore-required:minimized", result.RestoreSkipReason);
        }

        [Fact]
        public void Decide_WhenWindowIsMaximized_Restores()
        {
            ExcelWindowRestoreDecision result = ExcelWindowRestoreDecisionPolicy.Decide(
                windowResolved: true,
                visibleReadSucceeded: true,
                visible: true,
                placementReadSucceeded: true,
                showCmd: ExcelWindowRestoreDecisionPolicy.SwShowMaximized);

            Assert.True(result.ShouldRestore);
            Assert.Equal("False", result.RestoreSkipped);
            Assert.Equal("restore-required:maximized", result.RestoreSkipReason);
        }

        [Theory]
        [InlineData(false, true, true, true, ExcelWindowRestoreDecisionPolicy.SwShowNormal, "restore-required:window-null")]
        [InlineData(true, false, true, true, ExcelWindowRestoreDecisionPolicy.SwShowNormal, "restore-required:visible-read-failed")]
        [InlineData(true, true, false, true, ExcelWindowRestoreDecisionPolicy.SwShowNormal, "restore-required:not-visible")]
        [InlineData(true, true, true, false, ExcelWindowRestoreDecisionPolicy.SwShowNormal, "restore-required:placement-read-failed")]
        [InlineData(true, true, true, true, ExcelWindowRestoreDecisionPolicy.SwHide, "restore-required:hidden")]
        public void Decide_WhenRestoreIsRequired_ReturnsReason(
            bool windowResolved,
            bool visibleReadSucceeded,
            bool visible,
            bool placementReadSucceeded,
            int showCmd,
            string expectedReason)
        {
            ExcelWindowRestoreDecision result = ExcelWindowRestoreDecisionPolicy.Decide(
                windowResolved: windowResolved,
                visibleReadSucceeded: visibleReadSucceeded,
                visible: visible,
                placementReadSucceeded: placementReadSucceeded,
                showCmd: showCmd);

            Assert.True(result.ShouldRestore);
            Assert.Equal("False", result.RestoreSkipped);
            Assert.Equal(expectedReason, result.RestoreSkipReason);
        }
    }
}
