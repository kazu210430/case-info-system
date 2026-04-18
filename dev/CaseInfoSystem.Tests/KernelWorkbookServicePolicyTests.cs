using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookServicePolicyTests
    {
        [Fact]
        public void ResolveKernelWorkbookPath_ReturnsEmpty_WhenOpenKernelWorkbookAlreadyExists()
        {
            bool resolverCalled = false;

            string result = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: true,
                systemRoot: "C:\\案件",
                resolvePath: root =>
                {
                    resolverCalled = true;
                    return root + "\\案件情報System.xlsm";
                });

            Assert.False(resolverCalled);
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void ResolveKernelWorkbookPath_ReturnsEmpty_WhenSystemRootIsMissing()
        {
            string result = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: false,
                systemRoot: string.Empty,
                resolvePath: root => root + "\\案件情報System.xlsm");

            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void ResolveKernelWorkbookPath_UsesResolvedPath_WhenFallbackLookupIsNeeded()
        {
            string result = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: false,
                systemRoot: "C:\\案件",
                resolvePath: root => root + "\\案件情報System.xlsm");

            Assert.Equal("C:\\案件\\案件情報System.xlsm", result);
        }

        [Fact]
        public void Decide_ReturnsQuitFlow_WhenNoOtherWorkbookExists()
        {
            bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                saveKernelWorkbook: false,
                isKernelCaseCreationFlowActive: false,
                otherVisibleWorkbookExists: false,
                otherWorkbookExists: false);
            KernelHomeSessionCompletionAction action = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                skipDisplayRestoreForCaseCreation,
                otherVisibleWorkbookExists: false,
                otherWorkbookExists: false);

            Assert.False(skipDisplayRestoreForCaseCreation);
            Assert.Equal(
                KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithoutShowingExcelAndQuit,
                action);
        }

        [Fact]
        public void Decide_DismissesPreparedState_WhenCaseCreationShouldPreserveForegroundWorkbook()
        {
            bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                saveKernelWorkbook: true,
                isKernelCaseCreationFlowActive: true,
                otherVisibleWorkbookExists: true,
                otherWorkbookExists: true);
            KernelHomeSessionCompletionAction action = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                skipDisplayRestoreForCaseCreation,
                otherVisibleWorkbookExists: true,
                otherWorkbookExists: true);

            Assert.True(skipDisplayRestoreForCaseCreation);
            Assert.Equal(
                KernelHomeSessionCompletionAction.DismissPreparedHomeDisplayState,
                action);
        }

        [Fact]
        public void Decide_RestoresHomeDisplay_WhenOtherWorkbookExistsOutsideCaseCreationFlow()
        {
            bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                saveKernelWorkbook: true,
                isKernelCaseCreationFlowActive: false,
                otherVisibleWorkbookExists: true,
                otherWorkbookExists: true);
            KernelHomeSessionCompletionAction action = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                skipDisplayRestoreForCaseCreation,
                otherVisibleWorkbookExists: true,
                otherWorkbookExists: true);

            Assert.False(skipDisplayRestoreForCaseCreation);
            Assert.Equal(
                KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithShowingExcel,
                action);
        }
    }
}
