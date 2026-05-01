using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
        public void ResolveKernelWorkbookPathFromAvailableWorkbooks_UsesActiveKernelWorkbookDirectoryFallback()
        {
            string tempDirectory = Path.Combine(Path.GetTempPath(), "KernelWorkbookStateServiceTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDirectory);
            try
            {
                string kernelWorkbookPath = Path.Combine(tempDirectory, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"));
                File.WriteAllText(kernelWorkbookPath, string.Empty);

                var application = new Excel.Application();
                var workbook = new Excel.Workbook
                {
                    Application = application,
                    FullName = kernelWorkbookPath,
                    Name = Path.GetFileName(kernelWorkbookPath),
                    CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                };
                application.Workbooks.Add(workbook);
                application.ActiveWorkbook = workbook;

                var loggerMessages = new List<string>();
                var logger = OrchestrationTestSupport.CreateLogger(loggerMessages);
                var pathCompatibilityService = new PathCompatibilityService();
                var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);

                string resolved = KernelWorkbookResolver.ResolveKernelWorkbookPathFromAvailableWorkbooks(
                    application,
                    excelInteropService,
                    pathCompatibilityService,
                    logger,
                    workbookCandidate => WorkbookFileNameResolver.IsKernelWorkbookName(workbookCandidate == null ? string.Empty : workbookCandidate.Name));

                Assert.Equal(pathCompatibilityService.NormalizePath(kernelWorkbookPath), resolved);
            }
            finally
            {
                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, recursive: true);
                }
            }
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
