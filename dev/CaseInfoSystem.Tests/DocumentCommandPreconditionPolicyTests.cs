using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class DocumentCommandPreconditionPolicyTests
    {
        [Fact]
        public void Decide_ReturnsBlockBecauseIneligible_WhenEligibilityFailed()
        {
            DocumentCommandPreconditionDecision decision = DocumentCommandPreconditionPolicy.Decide(
                canExecuteInVsto: false,
                isVstoExecutionAllowed: true);

            Assert.Equal(DocumentCommandPreconditionDecision.BlockBecauseIneligible, decision);
        }

        [Fact]
        public void Decide_ReturnsBlockBecauseIneligible_WhenBothChecksFail()
        {
            DocumentCommandPreconditionDecision decision = DocumentCommandPreconditionPolicy.Decide(
                canExecuteInVsto: false,
                isVstoExecutionAllowed: false);

            Assert.Equal(DocumentCommandPreconditionDecision.BlockBecauseIneligible, decision);
        }

        [Fact]
        public void Decide_ReturnsBlockBecauseNotAllowlisted_WhenEligibilityPassedButAllowlistFailed()
        {
            DocumentCommandPreconditionDecision decision = DocumentCommandPreconditionPolicy.Decide(
                canExecuteInVsto: true,
                isVstoExecutionAllowed: false);

            Assert.Equal(DocumentCommandPreconditionDecision.BlockBecauseNotAllowlisted, decision);
        }

        [Fact]
        public void Decide_ReturnsContinue_WhenEligibilityAndAllowlistPass()
        {
            DocumentCommandPreconditionDecision decision = DocumentCommandPreconditionPolicy.Decide(
                canExecuteInVsto: true,
                isVstoExecutionAllowed: true);

            Assert.Equal(DocumentCommandPreconditionDecision.Continue, decision);
        }
    }
}
