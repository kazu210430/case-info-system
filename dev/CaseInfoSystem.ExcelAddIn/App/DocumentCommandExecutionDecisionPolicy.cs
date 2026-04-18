namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum DocumentCommandExecutionDecision
    {
        Continue,
        ThrowBecauseIneligible,
        ThrowBecauseNotAllowlisted
    }

    internal static class DocumentCommandExecutionDecisionPolicy
    {
        internal static DocumentCommandExecutionDecision Decide(DocumentCommandPreconditionDecision preconditionDecision)
        {
            if (preconditionDecision == DocumentCommandPreconditionDecision.Continue)
            {
                return DocumentCommandExecutionDecision.Continue;
            }

            if (preconditionDecision == DocumentCommandPreconditionDecision.BlockBecauseNotAllowlisted)
            {
                return DocumentCommandExecutionDecision.ThrowBecauseNotAllowlisted;
            }

            return DocumentCommandExecutionDecision.ThrowBecauseIneligible;
        }
    }
}
