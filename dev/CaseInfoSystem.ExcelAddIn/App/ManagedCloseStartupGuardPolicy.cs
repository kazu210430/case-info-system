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
        internal const string DefaultDelayReason = "defaultEligibleStartupGuard";
        internal const string GuardedRestoreEmptyStartupDelayReason = "guardedRestoreEmptyStartup";
        internal const string NotEligibleDelayReason = "notEligible";

        internal static ManagedCloseStartupGuardDelayDecision Decide(ManagedCloseStartupGuardFacts facts)
        {
            if (!IsEligible(facts))
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

            return new ManagedCloseStartupGuardDelayDecision(
                isEligible: true,
                delayMs: DefaultDelayMs,
                delayReason: DefaultDelayReason,
                usesGuardedRestoreEmptyStartupDelay: false);
        }

        internal static bool IsEligible(ManagedCloseStartupGuardFacts facts)
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

            return facts.CommandLineHasRestoreSwitch
                && !facts.CommandLineHasEmbeddingSwitch;
        }

        private static bool IsGuardedRestoreEmptyStartup(ManagedCloseStartupGuardFacts facts)
        {
            return facts != null
                && facts.ApplicationVisible
                && facts.CommandLineHasRestoreSwitch
                && !facts.CommandLineHasEmbeddingSwitch;
        }
    }
}
