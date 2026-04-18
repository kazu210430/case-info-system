using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class DocumentCommandExecutionDecisionPolicyTests
    {
        [Fact]
        public void Decide_ReturnsContinue_ForContinuePrecondition()
        {
            DocumentCommandExecutionDecision decision = DocumentCommandExecutionDecisionPolicy.Decide(
                DocumentCommandPreconditionDecision.Continue);

            Assert.Equal(DocumentCommandExecutionDecision.Continue, decision);
        }

        [Fact]
        public void Decide_ReturnsThrowBecauseIneligible_ForIneligiblePrecondition()
        {
            DocumentCommandExecutionDecision decision = DocumentCommandExecutionDecisionPolicy.Decide(
                DocumentCommandPreconditionDecision.BlockBecauseIneligible);

            Assert.Equal(DocumentCommandExecutionDecision.ThrowBecauseIneligible, decision);
        }

        [Fact]
        public void Decide_ReturnsThrowBecauseNotAllowlisted_ForNotAllowlistedPrecondition()
        {
            DocumentCommandExecutionDecision decision = DocumentCommandExecutionDecisionPolicy.Decide(
                DocumentCommandPreconditionDecision.BlockBecauseNotAllowlisted);

            Assert.Equal(DocumentCommandExecutionDecision.ThrowBecauseNotAllowlisted, decision);
        }

        [Fact]
        public void Decide_ReturnsThrowBecauseIneligible_ForUnknownPreconditionValue()
        {
            DocumentCommandExecutionDecision decision = DocumentCommandExecutionDecisionPolicy.Decide(
                (DocumentCommandPreconditionDecision)999);

            Assert.Equal(DocumentCommandExecutionDecision.ThrowBecauseIneligible, decision);
        }
    }
}
