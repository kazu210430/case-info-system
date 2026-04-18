namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum DocumentCommandPreconditionDecision
    {
        Continue,
        BlockBecauseIneligible,
        BlockBecauseNotAllowlisted
    }

    internal static class DocumentCommandPreconditionPolicy
    {
        internal static DocumentCommandPreconditionDecision Decide(
            bool canExecuteInVsto,
            bool isVstoExecutionAllowed)
        {
            if (!canExecuteInVsto)
            {
                return DocumentCommandPreconditionDecision.BlockBecauseIneligible;
            }

            if (!isVstoExecutionAllowed)
            {
                return DocumentCommandPreconditionDecision.BlockBecauseNotAllowlisted;
            }

            return DocumentCommandPreconditionDecision.Continue;
        }
    }
}
