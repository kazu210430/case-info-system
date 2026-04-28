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
                canExecuteInVsto: false);

            Assert.Equal(DocumentCommandPreconditionDecision.BlockBecauseIneligible, decision);
        }

        [Fact]
        public void Decide_ReturnsContinue_WhenEligibilityPasses()
        {
            DocumentCommandPreconditionDecision decision = DocumentCommandPreconditionPolicy.Decide(
                canExecuteInVsto: true);

            Assert.Equal(DocumentCommandPreconditionDecision.Continue, decision);
        }
    }
}
