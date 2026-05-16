using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class CreatedCaseDisplaySessionStateReaderTests
    {
        [Fact]
        public void DecideStart_AllowsOnlyCreatedCaseReasonWithWorkbookFullName()
        {
            var reader = new CreatedCaseDisplaySessionStateReader();

            CreatedCaseDisplaySessionStartDecision allowed = reader.DecideStart(
                new CreatedCaseDisplaySessionStartInput(
                    isCreatedCaseDisplayReason: true,
                    workbookFullName: @"C:\cases\case.xlsx"));
            CreatedCaseDisplaySessionStartDecision wrongReason = reader.DecideStart(
                new CreatedCaseDisplaySessionStartInput(
                    isCreatedCaseDisplayReason: false,
                    workbookFullName: @"C:\cases\case.xlsx"));
            CreatedCaseDisplaySessionStartDecision missingWorkbook = reader.DecideStart(
                new CreatedCaseDisplaySessionStartInput(
                    isCreatedCaseDisplayReason: true,
                    workbookFullName: string.Empty));

            Assert.True(allowed.ShouldStart);
            Assert.Equal(string.Empty, allowed.BlockedReason);
            Assert.False(wrongReason.ShouldStart);
            Assert.Equal("reasonNotCreatedCaseDisplay", wrongReason.BlockedReason);
            Assert.False(missingWorkbook.ShouldStart);
            Assert.Equal("workbookFullName=nullOrEmpty", missingWorkbook.BlockedReason);
        }

        [Fact]
        public void ResolveForCompletion_UsesExactWorkbookBeforeSingleActiveFallback()
        {
            var reader = new CreatedCaseDisplaySessionStateReader();
            var first = new CreatedCaseDisplaySessionSnapshot(
                "CDS-0001",
                @"C:\cases\first.xlsx",
                "created",
                isCompleted: false);
            var second = new CreatedCaseDisplaySessionSnapshot(
                "CDS-0002",
                @"C:\cases\second.xlsx",
                "created",
                isCompleted: false);

            CreatedCaseDisplaySessionSnapshot exact = reader.ResolveForCompletion(
                new CreatedCaseDisplaySessionResolutionInput(
                    isCreatedCaseDisplayReason: true,
                    workbookFullName: @"C:\CASES\SECOND.xlsx",
                    activeSessions: new[] { first, second }));
            CreatedCaseDisplaySessionSnapshot ambiguous = reader.ResolveForCompletion(
                new CreatedCaseDisplaySessionResolutionInput(
                    isCreatedCaseDisplayReason: true,
                    workbookFullName: @"C:\cases\missing.xlsx",
                    activeSessions: new[] { first, second }));
            CreatedCaseDisplaySessionSnapshot singleFallback = reader.ResolveForCompletion(
                new CreatedCaseDisplaySessionResolutionInput(
                    isCreatedCaseDisplayReason: true,
                    workbookFullName: string.Empty,
                    activeSessions: new[] { first }));

            Assert.Same(second, exact);
            Assert.Null(ambiguous);
            Assert.Same(first, singleFallback);
        }

        [Fact]
        public void Source_DoesNotOwnEmitMutationRetryTimerOrForegroundExecution()
        {
            string source = ReadAppSource("CreatedCaseDisplaySessionStateReader.cs");

            Assert.Contains("ResolveForCompletion", source);
            Assert.DoesNotContain("_createdCaseDisplaySessions", source);
            Assert.DoesNotContain("Remove(", source);
            Assert.DoesNotContain(".IsCompleted =", source);
            Assert.DoesNotContain("NewCaseVisibilityObservation", source);
            Assert.DoesNotContain("case-display-completed", source);
            Assert.DoesNotContain("TaskPaneRetryTimerLifecycle", source);
            Assert.DoesNotContain("ExecuteFinalForegroundGuaranteeRecovery", source);
        }

        private static string ReadAppSource(string appFileName)
        {
            string repoRoot = FindRepositoryRoot();
            return File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", appFileName));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
