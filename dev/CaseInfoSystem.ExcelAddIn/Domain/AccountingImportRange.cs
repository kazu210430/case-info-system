namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class AccountingImportRange
    {
        /// <summary>
        internal AccountingImportRange(int startRound, int endRound)
        {
            StartRound = startRound;
            EndRound = endRound;
        }

        internal int StartRound { get; }

        internal int EndRound { get; }
    }
}
