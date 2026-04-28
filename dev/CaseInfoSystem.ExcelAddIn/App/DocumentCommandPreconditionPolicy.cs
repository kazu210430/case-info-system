namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum DocumentCommandPreconditionDecision
    {
        Continue,
        BlockBecauseIneligible
    }

    internal static class DocumentCommandPreconditionPolicy
    {
        internal static DocumentCommandPreconditionDecision Decide(bool canExecuteInVsto)
        {
            if (!canExecuteInVsto)
            {
                return DocumentCommandPreconditionDecision.BlockBecauseIneligible;
            }

            return DocumentCommandPreconditionDecision.Continue;
        }
    }
}
