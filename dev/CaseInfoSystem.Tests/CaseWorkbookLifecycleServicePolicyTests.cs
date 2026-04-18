using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class CaseWorkbookLifecycleServicePolicyTests
    {
        [Fact]
        public void DecideInitialization_ReturnsNone_WhenWorkbookIsOutOfScope()
        {
            CaseWorkbookInitializationAction action = CaseWorkbookLifecycleInitializationPolicy.Decide(
                isBaseOrCaseWorkbook: false,
                workbookKey: "C:\\case.xlsx",
                isAlreadyInitialized: false,
                isCaseWorkbook: true);

            Assert.Equal(CaseWorkbookInitializationAction.None, action);
        }

        [Fact]
        public void DecideInitialization_ReturnsNone_WhenWorkbookKeyIsMissing()
        {
            CaseWorkbookInitializationAction action = CaseWorkbookLifecycleInitializationPolicy.Decide(
                isBaseOrCaseWorkbook: true,
                workbookKey: string.Empty,
                isAlreadyInitialized: false,
                isCaseWorkbook: true);

            Assert.Equal(CaseWorkbookInitializationAction.None, action);
        }

        [Fact]
        public void DecideInitialization_ReturnsNone_WhenWorkbookWasAlreadyInitialized()
        {
            CaseWorkbookInitializationAction action = CaseWorkbookLifecycleInitializationPolicy.Decide(
                isBaseOrCaseWorkbook: true,
                workbookKey: "C:\\case.xlsx",
                isAlreadyInitialized: true,
                isCaseWorkbook: true);

            Assert.Equal(CaseWorkbookInitializationAction.None, action);
        }

        [Fact]
        public void DecideInitialization_ReturnsInitializeCaseWorkbook_ForNewCaseWorkbook()
        {
            CaseWorkbookInitializationAction action = CaseWorkbookLifecycleInitializationPolicy.Decide(
                isBaseOrCaseWorkbook: true,
                workbookKey: "C:\\case.xlsx",
                isAlreadyInitialized: false,
                isCaseWorkbook: true);

            Assert.Equal(CaseWorkbookInitializationAction.InitializeCaseWorkbook, action);
        }

        [Fact]
        public void DecideInitialization_ReturnsInitializeBaseWorkbook_ForNewBaseWorkbook()
        {
            CaseWorkbookInitializationAction action = CaseWorkbookLifecycleInitializationPolicy.Decide(
                isBaseOrCaseWorkbook: true,
                workbookKey: "C:\\base.xlsm",
                isAlreadyInitialized: false,
                isCaseWorkbook: false);

            Assert.Equal(CaseWorkbookInitializationAction.InitializeBaseWorkbook, action);
        }

        [Fact]
        public void DecideBeforeClose_ReturnsIgnore_WhenWorkbookIsOutOfScope()
        {
            CaseWorkbookBeforeCloseAction action = CaseWorkbookBeforeClosePolicy.Decide(
                isBaseOrCaseWorkbook: false,
                isManagedClose: false,
                isSessionDirty: true);

            Assert.Equal(CaseWorkbookBeforeCloseAction.Ignore, action);
        }

        [Fact]
        public void DecideBeforeClose_ReturnsSuppressPromptForManagedClose_WhenManagedCloseIsActive()
        {
            CaseWorkbookBeforeCloseAction action = CaseWorkbookBeforeClosePolicy.Decide(
                isBaseOrCaseWorkbook: true,
                isManagedClose: true,
                isSessionDirty: true);

            Assert.Equal(CaseWorkbookBeforeCloseAction.SuppressPromptForManagedClose, action);
        }

        [Fact]
        public void DecideBeforeClose_ReturnsPromptForDirtySession_WhenWorkbookWasEdited()
        {
            CaseWorkbookBeforeCloseAction action = CaseWorkbookBeforeClosePolicy.Decide(
                isBaseOrCaseWorkbook: true,
                isManagedClose: false,
                isSessionDirty: true);

            Assert.Equal(CaseWorkbookBeforeCloseAction.PromptForDirtySession, action);
        }

        [Fact]
        public void DecideBeforeClose_ReturnsSchedulePostCloseFollowUp_WhenWorkbookIsClean()
        {
            CaseWorkbookBeforeCloseAction action = CaseWorkbookBeforeClosePolicy.Decide(
                isBaseOrCaseWorkbook: true,
                isManagedClose: false,
                isSessionDirty: false);

            Assert.Equal(CaseWorkbookBeforeCloseAction.SchedulePostCloseFollowUp, action);
        }

        [Theory]
        [InlineData(false, false, false, 0)]
        [InlineData(true, true, false, 0)]
        [InlineData(true, false, true, 1)]
        [InlineData(true, false, false, 2)]
        public void DecideSheetChange_UsesWorkbookScopeManagedCloseAndSuppression(
            bool isBaseOrCaseWorkbook,
            bool isManagedClose,
            bool isTransientPaneSuppressed,
            int expected)
        {
            CaseWorkbookSheetChangeAction result = CaseWorkbookSheetChangePolicy.Decide(
                isBaseOrCaseWorkbook,
                isManagedClose,
                isTransientPaneSuppressed);

            Assert.Equal((CaseWorkbookSheetChangeAction)expected, result);
        }
    }
}
