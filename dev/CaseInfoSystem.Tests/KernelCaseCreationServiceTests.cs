using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class KernelCaseCreationServiceTests
    {
        [Fact]
        public void CreateSavedCase_UsesLocalWorkingCopyForSyncRootAndFinalizesBackToFinalPath()
        {
            List<string> logs = new List<string>();
            List<string> openedPaths = new List<string>();
            List<string> initializedPaths = new List<string>();
            string tempRoot = Path.Combine(Path.GetTempPath(), "KernelCaseCreationServiceTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            string syncRoot = Path.Combine(tempRoot, "OneDrive");
            string caseFolderPath = Path.Combine(syncRoot, "ClientA");
            Directory.CreateDirectory(caseFolderPath);
            string finalCaseWorkbookPath = Path.Combine(caseFolderPath, "20260420_CASE.xlsx");
            File.WriteAllText(finalCaseWorkbookPath, "seed");

            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\System\案件情報System.xlsm",
                Name = "案件情報System.xlsm",
                Path = @"C:\System"
            };

            Excel.Workbook hiddenWorkbook = new Excel.Workbook
            {
                FullName = finalCaseWorkbookPath,
                Name = "20260420_CASE.xlsx",
                Path = caseFolderPath
            };

            Excel.Application hiddenApplication = new Excel.Application();
            string localWorkingCaseWorkbookPath = Path.Combine(tempRoot, "temp", "20260420_CASE.xlsx");
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => kernelWorkbook
                });

            var caseWorkbookOpenStrategy = new CaseWorkbookOpenStrategy
            {
                OnOpenHiddenWorkbook = caseWorkbookPath =>
                {
                    openedPaths.Add(caseWorkbookPath ?? string.Empty);
                    hiddenWorkbook.FullName = caseWorkbookPath ?? string.Empty;
                    hiddenWorkbook.Name = Path.GetFileName(caseWorkbookPath ?? string.Empty) ?? string.Empty;
                    hiddenWorkbook.Path = Path.GetDirectoryName(caseWorkbookPath ?? string.Empty) ?? string.Empty;
                    return new CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession(hiddenApplication, hiddenWorkbook);
                }
            };

            var kernelCasePathService = new KernelCasePathService
            {
                OnIsUnderSyncRoot = path => string.Equals(path, finalCaseWorkbookPath, StringComparison.OrdinalIgnoreCase),
                OnBuildLocalWorkingCaseWorkbookPath = path =>
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(localWorkingCaseWorkbookPath) ?? tempRoot);
                    return localWorkingCaseWorkbookPath;
                },
                OnMoveLocalWorkingCaseToFinalPath = (localPath, finalPath) =>
                {
                    File.Copy(localPath, finalPath, overwrite: true);
                    File.Delete(localPath);
                    return true;
                }
            };

            var caseWorkbookInitializer = new CaseWorkbookInitializer
            {
                OnInitializeForHiddenCreate = (sourceKernelWorkbook, createdCaseWorkbook, plan) =>
                {
                    initializedPaths.Add(createdCaseWorkbook == null ? string.Empty : createdCaseWorkbook.FullName ?? string.Empty);
                }
            };

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                kernelCasePathService,
                caseWorkbookInitializer,
                caseWorkbookOpenStrategy,
                new FolderWindowService(),
                new TransientPaneSuppressionService(),
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var plan = new KernelCaseCreationPlan
            {
                Mode = KernelCaseCreationMode.NewCaseDefault,
                CustomerName = string.Empty,
                SystemRoot = @"C:\System",
                BaseWorkbookPath = @"C:\System\BASE.xlsx",
                CaseFolderPath = caseFolderPath,
                CaseWorkbookPath = finalCaseWorkbookPath,
                NameRuleA = "YYYY",
                NameRuleB = "DOC"
            };

            try
            {
                KernelCaseCreationResult result = service.CreateSavedCase(kernelWorkbook, plan);

                Assert.Equal(new[] { localWorkingCaseWorkbookPath }, openedPaths);
                Assert.Equal(new[] { localWorkingCaseWorkbookPath }, initializedPaths);
                Assert.Equal(1, hiddenWorkbook.SaveCallCount);
                Assert.True(result.Success);
                Assert.Equal(finalCaseWorkbookPath, result.CaseWorkbookPath);
                Assert.True(File.Exists(finalCaseWorkbookPath));
                Assert.False(File.Exists(localWorkingCaseWorkbookPath));
                Assert.Contains(logs, message => message.IndexOf("local working copy prepared", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.Contains(logs, message => message.IndexOf("local working copy finalized", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            finally
            {
                if (Directory.Exists(tempRoot))
                {
                    Directory.Delete(tempRoot, recursive: true);
                }
            }
        }
    }
}
