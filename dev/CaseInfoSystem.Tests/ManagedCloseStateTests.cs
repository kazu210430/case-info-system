using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ManagedCloseStateTests
    {
        [Fact]
        public void BeginScope_MarksWorkbookAsManagedClose_UntilScopeIsDisposed()
        {
            var state = new ManagedCloseState();

            using (state.BeginScope("case-1"))
            {
                Assert.True(state.IsManagedClose("case-1"));
            }

            Assert.False(state.IsManagedClose("case-1"));
        }

        [Fact]
        public void BeginScope_KeepsManagedCloseActive_UntilOutermostScopeIsDisposed()
        {
            var state = new ManagedCloseState();

            using (state.BeginScope("case-1"))
            {
                using (state.BeginScope("case-1"))
                {
                    Assert.True(state.IsManagedClose("case-1"));
                }

                Assert.True(state.IsManagedClose("case-1"));
            }

            Assert.False(state.IsManagedClose("case-1"));
        }
    }
}
