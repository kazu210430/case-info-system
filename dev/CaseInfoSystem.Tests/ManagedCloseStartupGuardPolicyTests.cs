using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ManagedCloseStartupGuardPolicyTests
    {
        [Fact]
        public void Decide_UsesShortDelay_ForGuardedRestoreEmptyStartup()
        {
            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(
                CreateEligibleFacts(applicationVisible: true, commandLineHasRestoreSwitch: true));

            Assert.True(decision.IsEligible);
            Assert.Equal(ManagedCloseStartupGuardPolicy.GuardedRestoreEmptyStartupDelayMs, decision.DelayMs);
            Assert.True(decision.UsesGuardedRestoreEmptyStartupDelay);
            Assert.Equal(ManagedCloseStartupGuardPolicy.GuardedRestoreEmptyStartupDelayReason, decision.DelayReason);
        }

        [Fact]
        public void Decide_UsesDefaultDelay_ForInvisibleEligibleStartup()
        {
            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(
                CreateEligibleFacts(applicationVisible: false, commandLineHasRestoreSwitch: false));

            Assert.True(decision.IsEligible);
            Assert.Equal(ManagedCloseStartupGuardPolicy.DefaultDelayMs, decision.DelayMs);
            Assert.False(decision.UsesGuardedRestoreEmptyStartupDelay);
            Assert.Equal(ManagedCloseStartupGuardPolicy.DefaultDelayReason, decision.DelayReason);
        }

        [Fact]
        public void Decide_ReturnsIneligible_WhenReadFailed()
        {
            ManagedCloseStartupGuardFacts facts = CreateEligibleFacts(applicationVisible: true, commandLineHasRestoreSwitch: true);
            facts.ReadFailed = true;

            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(facts);

            Assert.False(decision.IsEligible);
            Assert.Equal(ManagedCloseStartupGuardPolicy.NotEligibleDelayReason, decision.DelayReason);
        }

        [Fact]
        public void Decide_ReturnsIneligible_ForVisibleStartupWithoutRestoreSwitch()
        {
            ManagedCloseStartupGuardFacts facts = CreateEligibleFacts(applicationVisible: true, commandLineHasRestoreSwitch: false);

            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(facts);

            Assert.False(decision.IsEligible);
            Assert.False(decision.UsesGuardedRestoreEmptyStartupDelay);
        }

        [Fact]
        public void Decide_ReturnsIneligible_ForVisibleRestoreStartupWithEmbeddingSwitch()
        {
            ManagedCloseStartupGuardFacts facts = CreateEligibleFacts(applicationVisible: true, commandLineHasRestoreSwitch: true);
            facts.CommandLineHasEmbeddingSwitch = true;

            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(facts);

            Assert.False(decision.IsEligible);
            Assert.False(decision.UsesGuardedRestoreEmptyStartupDelay);
        }

        [Theory]
        [InlineData(false, true, 0, false, false)]
        [InlineData(false, false, 1, false, false)]
        [InlineData(false, false, 0, true, false)]
        [InlineData(false, false, 0, false, true)]
        [InlineData(true, false, 0, false, false)]
        public void Decide_ReturnsIneligible_WhenWorkbookOrKernelSignalsExist(
            bool workbookOpenObserved,
            bool activeWorkbookPresent,
            int workbooksCount,
            bool visibleNonKernelWorkbookExists,
            bool hasOpenKernelWorkbook)
        {
            ManagedCloseStartupGuardFacts facts = CreateEligibleFacts(applicationVisible: true, commandLineHasRestoreSwitch: true);
            facts.WorkbookOpenObserved = workbookOpenObserved;
            facts.ActiveWorkbookPresent = activeWorkbookPresent;
            facts.WorkbooksCount = workbooksCount;
            facts.VisibleNonKernelWorkbookExists = visibleNonKernelWorkbookExists;
            facts.HasOpenKernelWorkbook = hasOpenKernelWorkbook;

            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(facts);

            Assert.False(decision.IsEligible);
            Assert.False(decision.UsesGuardedRestoreEmptyStartupDelay);
        }

        [Fact]
        public void Decide_ReturnsIneligible_WhenDelayedFactsDiscoverWorkbook()
        {
            ManagedCloseStartupGuardFacts facts = CreateEligibleFacts(applicationVisible: true, commandLineHasRestoreSwitch: true);
            facts.WorkbooksCount = 1;

            ManagedCloseStartupGuardDelayDecision decision = ManagedCloseStartupGuardPolicy.Decide(facts);

            Assert.False(decision.IsEligible);
        }

        private static ManagedCloseStartupGuardFacts CreateEligibleFacts(
            bool applicationVisible,
            bool commandLineHasRestoreSwitch)
        {
            return new ManagedCloseStartupGuardFacts
            {
                ApplicationVisible = applicationVisible,
                CommandLineHasRestoreSwitch = commandLineHasRestoreSwitch,
                CommandLineHasEmbeddingSwitch = false,
                WorkbooksCount = 0,
                ActiveWorkbookPresent = false,
                VisibleNonKernelWorkbookExists = false,
                HasOpenKernelWorkbook = false,
                WorkbookOpenObserved = false,
                ReadFailed = false
            };
        }
    }
}
