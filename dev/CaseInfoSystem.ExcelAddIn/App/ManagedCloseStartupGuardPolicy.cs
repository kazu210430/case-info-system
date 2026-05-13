namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class ManagedCloseStartupGuardFacts
    {
        internal bool ReadFailed { get; set; }

        internal bool WorkbookOpenObserved { get; set; }

        internal bool ActiveWorkbookPresent { get; set; }

        internal int WorkbooksCount { get; set; }

        internal bool VisibleNonKernelWorkbookExists { get; set; }

        internal bool HasOpenKernelWorkbook { get; set; }

        internal bool ApplicationVisible { get; set; }

        internal bool CommandLineHasRestoreSwitch { get; set; }

        internal bool CommandLineHasEmbeddingSwitch { get; set; }

        internal bool CommandLineHasWorkbookFileArgument { get; set; }

        internal int ParentProcessId { get; set; }

        internal bool ApplicationUserControl { get; set; }

        internal bool ApplicationUserControlReadFailed { get; set; }
    }

    internal sealed class ManagedCloseStartupGuardMarkerFacts
    {
        internal ManagedWorkbookCloseMarkerKind Kind { get; set; }

        internal int WriterProcessId { get; set; }
    }

    internal sealed class ManagedCloseStartupGuardDelayDecision
    {
        internal ManagedCloseStartupGuardDelayDecision(
            bool isEligible,
            int delayMs,
            string delayReason,
            bool usesGuardedRestoreEmptyStartupDelay)
        {
            IsEligible = isEligible;
            DelayMs = delayMs;
            DelayReason = delayReason;
            UsesGuardedRestoreEmptyStartupDelay = usesGuardedRestoreEmptyStartupDelay;
        }

        internal bool IsEligible { get; }

        internal int DelayMs { get; }

        internal string DelayReason { get; }

        internal bool UsesGuardedRestoreEmptyStartupDelay { get; }
    }

    internal static class ManagedCloseStartupGuardPolicy
    {
        internal const int DefaultDelayMs = 1000;
        internal const int GuardedRestoreEmptyStartupDelayMs = 250;
        internal const int AccountingCloseOwnedEmptyStartupDelayMs = 50;
        internal const string DefaultDelayReason = "defaultEligibleStartupGuard";
        internal const string GuardedRestoreEmptyStartupDelayReason = "guardedRestoreEmptyStartup";
        internal const string AccountingCloseOwnedEmptyStartupDelayReason = "accountingCloseOwnedEmptyStartup";
        internal const string NotEligibleDelayReason = "notEligible";

        internal static ManagedCloseStartupGuardDelayDecision Decide(ManagedCloseStartupGuardFacts facts)
        {
            return Decide(facts, null);
        }

        internal static ManagedCloseStartupGuardDelayDecision Decide(
            ManagedCloseStartupGuardFacts facts,
            ManagedCloseStartupGuardMarkerFacts markerFacts)
        {
            if (!IsEligible(facts, markerFacts))
            {
                return new ManagedCloseStartupGuardDelayDecision(
                    isEligible: false,
                    delayMs: DefaultDelayMs,
                    delayReason: NotEligibleDelayReason,
                    usesGuardedRestoreEmptyStartupDelay: false);
            }

            if (IsGuardedRestoreEmptyStartup(facts))
            {
                return new ManagedCloseStartupGuardDelayDecision(
                    isEligible: true,
                    delayMs: GuardedRestoreEmptyStartupDelayMs,
                    delayReason: GuardedRestoreEmptyStartupDelayReason,
                    usesGuardedRestoreEmptyStartupDelay: true);
            }

            if (IsAccountingCloseOwnedEmptyStartup(facts, markerFacts))
            {
                return new ManagedCloseStartupGuardDelayDecision(
                    isEligible: true,
                    delayMs: AccountingCloseOwnedEmptyStartupDelayMs,
                    delayReason: AccountingCloseOwnedEmptyStartupDelayReason,
                    usesGuardedRestoreEmptyStartupDelay: false);
            }

            return new ManagedCloseStartupGuardDelayDecision(
                isEligible: true,
                delayMs: DefaultDelayMs,
                delayReason: DefaultDelayReason,
                usesGuardedRestoreEmptyStartupDelay: false);
        }

        internal static bool IsEligible(ManagedCloseStartupGuardFacts facts)
        {
            return IsEligible(facts, null);
        }

        internal static bool IsEligible(ManagedCloseStartupGuardFacts facts, ManagedCloseStartupGuardMarkerFacts markerFacts)
        {
            if (facts == null
                || facts.ReadFailed
                || facts.WorkbookOpenObserved
                || facts.ActiveWorkbookPresent
                || facts.WorkbooksCount != 0
                || facts.VisibleNonKernelWorkbookExists
                || facts.HasOpenKernelWorkbook)
            {
                return false;
            }

            if (!facts.ApplicationVisible)
            {
                return true;
            }

            return (facts.CommandLineHasRestoreSwitch
                && !facts.CommandLineHasEmbeddingSwitch)
                || IsAccountingCloseOwnedEmptyStartup(facts, markerFacts);
        }

        private static bool IsGuardedRestoreEmptyStartup(ManagedCloseStartupGuardFacts facts)
        {
            return facts != null
                && facts.ApplicationVisible
                && facts.CommandLineHasRestoreSwitch
                && !facts.CommandLineHasEmbeddingSwitch;
        }

        private static bool IsAccountingCloseOwnedEmptyStartup(
            ManagedCloseStartupGuardFacts facts,
            ManagedCloseStartupGuardMarkerFacts markerFacts)
        {
            return facts != null
                && markerFacts != null
                && facts.ApplicationVisible
                && !facts.CommandLineHasRestoreSwitch
                && !facts.CommandLineHasEmbeddingSwitch
                && markerFacts.Kind == ManagedWorkbookCloseMarkerKind.AccountingClose
                && markerFacts.WriterProcessId > 0
                && (facts.ParentProcessId == markerFacts.WriterProcessId
                    || IsAccountingCloseAutomationEmptyStartup(facts));
        }

        private static bool IsAccountingCloseAutomationEmptyStartup(ManagedCloseStartupGuardFacts facts)
        {
            return facts != null
                && !facts.ApplicationUserControlReadFailed
                && !facts.ApplicationUserControl
                && !facts.CommandLineHasWorkbookFileArgument;
        }
    }
}
