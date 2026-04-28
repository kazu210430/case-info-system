namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum DocumentCommandExecutionDecision
    {
        Continue,
        ThrowBecauseIneligible
    }

    internal static class DocumentCommandExecutionDecisionPolicy
    {
        internal static DocumentCommandExecutionDecision Decide(DocumentCommandPreconditionDecision preconditionDecision)
        {
            if (preconditionDecision == DocumentCommandPreconditionDecision.Continue)
            {
                return DocumentCommandExecutionDecision.Continue;
            }

            return DocumentCommandExecutionDecision.ThrowBecauseIneligible;
        }
    }
}
