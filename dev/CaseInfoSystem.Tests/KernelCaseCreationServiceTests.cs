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
        public void CreateSavedCase_OpensWorkbookInCurrentApplicationAndClosesAfterSave()
        {
            List<string> logs = new List<string>();
            List<string> initializedPaths = new List<string>();
            string tempRoot = Path.Combine(Path.GetTempPath(), "KernelCaseCreationServiceTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            string caseFolderPath = Path.Combine(tempRoot, "ClientA");
            Directory.CreateDirectory(caseFolderPath);
            string finalCaseWorkbookPath = Path.Combine(caseFolderPath, "20260420_CASE.xlsx");
            File.WriteAllText(finalCaseWorkbookPath, "seed");

            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\System\案件情報System.xlsm",
                Name = "案件情報System.xlsm",
                Path = @"C:\System"
            };
            kernelWorkbook.Application = new Excel.Application();
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => kernelWorkbook
                });

            var caseWorkbookInitializer = new CaseWorkbookInitializer
            {
                OnInitializeForVisibleCreate = (sourceKernelWorkbook, createdCaseWorkbook, plan) =>
                {
                    initializedPaths.Add(createdCaseWorkbook == null ? string.Empty : createdCaseWorkbook.FullName ?? string.Empty);
                },
                OnInitializeForHiddenCreate = (sourceKernelWorkbook, createdCaseWorkbook, plan) =>
                {
                    initializedPaths.Add(createdCaseWorkbook == null ? string.Empty : createdCaseWorkbook.FullName ?? string.Empty);
                }
            };

            var transientPaneSuppressionService = new TransientPaneSuppressionService(
                new FakeExcelInteropService(),
                new PathCompatibilityService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var caseWorkbookLifecycleService = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(logs),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = workbook => workbook == null
                        ? string.Empty
                        : (workbook.FullName ?? workbook.Name ?? string.Empty)
                });

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                caseWorkbookInitializer,
                new CaseWorkbookOpenStrategy(),
                transientPaneSuppressionService,
                caseWorkbookLifecycleService,
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
                Excel.Workbook openedWorkbook = Assert.Single(kernelWorkbook.Application.Workbooks);

                Assert.Equal(new[] { finalCaseWorkbookPath }, initializedPaths);
                Assert.Equal(finalCaseWorkbookPath, openedWorkbook.FullName);
                Assert.Equal(1, openedWorkbook.SaveCallCount);
                Assert.Equal(1, openedWorkbook.CloseCallCount);
                Assert.True(result.Success);
                Assert.Equal(finalCaseWorkbookPath, result.CaseWorkbookPath);
                Assert.True(File.Exists(finalCaseWorkbookPath));
                Assert.Contains(logs, message => message.IndexOf("workbook opened", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.Contains(logs, message => message.IndexOf("workbook closed", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            finally
            {
                if (Directory.Exists(tempRoot))
                {
                    Directory.Delete(tempRoot, recursive: true);
                }
            }
        }

        [Fact]
        public void CreateSavedCase_RestoresDisplayAlertsOnSuccess()
        {
            List<string> logs = new List<string>();
            string tempRoot = Path.Combine(Path.GetTempPath(), "KernelCaseCreationServiceTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            string caseWorkbookPath = Path.Combine(tempRoot, "20260420_CASE.xlsx");
            File.WriteAllText(caseWorkbookPath, "seed");

            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\System\Kernel.xlsm",
                Name = "Kernel.xlsm",
                Path = @"C:\System"
            };
            kernelWorkbook.Application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => kernelWorkbook
                });

            var transientPaneSuppressionService = new TransientPaneSuppressionService(
                new FakeExcelInteropService(),
                new PathCompatibilityService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var caseWorkbookLifecycleService = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(logs),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = workbook => workbook == null
                        ? string.Empty
                        : (workbook.FullName ?? workbook.Name ?? string.Empty)
                });

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                new CaseWorkbookInitializer(),
                new CaseWorkbookOpenStrategy(),
                transientPaneSuppressionService,
                caseWorkbookLifecycleService,
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var plan = new KernelCaseCreationPlan
            {
                Mode = KernelCaseCreationMode.NewCaseDefault,
                CustomerName = string.Empty,
                SystemRoot = @"C:\System",
                BaseWorkbookPath = @"C:\System\BASE.xlsx",
                CaseFolderPath = tempRoot,
                CaseWorkbookPath = caseWorkbookPath,
                NameRuleA = "YYYY",
                NameRuleB = "DOC"
            };

            try
            {
                KernelCaseCreationResult result = service.CreateSavedCase(kernelWorkbook, plan);
                Excel.Workbook openedWorkbook = Assert.Single(kernelWorkbook.Application.Workbooks);

                Assert.True(result.Success);
                Assert.True(kernelWorkbook.Application.DisplayAlerts);
                Assert.Equal(1, openedWorkbook.CloseCallCount);
            }
            finally
            {
                if (Directory.Exists(tempRoot))
                {
                    Directory.Delete(tempRoot, recursive: true);
                }
            }
        }

        [Fact]
        public void CreateSavedCase_WhenInitializationFails_ClosesWorkbookAndRestoresDisplayAlerts()
        {
            List<string> logs = new List<string>();
            string tempRoot = Path.Combine(Path.GetTempPath(), "KernelCaseCreationServiceTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            string caseWorkbookPath = Path.Combine(tempRoot, "20260420_CASE.xlsx");
            File.WriteAllText(caseWorkbookPath, "seed");

            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\System\Kernel.xlsm",
                Name = "Kernel.xlsm",
                Path = @"C:\System"
            };
            kernelWorkbook.Application = new Excel.Application
            {
                DisplayAlerts = true
            };
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                    GetOpenKernelWorkbook = () => kernelWorkbook
                });

            var caseWorkbookInitializer = new CaseWorkbookInitializer
            {
                OnInitializeForVisibleCreate = (_, _, _) => throw new InvalidOperationException("boom"),
                OnInitializeForHiddenCreate = (_, _, _) => throw new InvalidOperationException("boom")
            };

            var transientPaneSuppressionService = new TransientPaneSuppressionService(
                new FakeExcelInteropService(),
                new PathCompatibilityService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var caseWorkbookLifecycleService = new CaseWorkbookLifecycleService(
                OrchestrationTestSupport.CreateLogger(logs),
                new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks
                {
                    GetWorkbookKey = workbook => workbook == null
                        ? string.Empty
                        : (workbook.FullName ?? workbook.Name ?? string.Empty)
                });

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                caseWorkbookInitializer,
                new CaseWorkbookOpenStrategy(),
                transientPaneSuppressionService,
                caseWorkbookLifecycleService,
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var plan = new KernelCaseCreationPlan
            {
                Mode = KernelCaseCreationMode.NewCaseDefault,
                CustomerName = string.Empty,
                SystemRoot = @"C:\System",
                BaseWorkbookPath = @"C:\System\BASE.xlsx",
                CaseFolderPath = tempRoot,
                CaseWorkbookPath = caseWorkbookPath,
                NameRuleA = "YYYY",
                NameRuleB = "DOC"
            };

            try
            {
                Assert.Throws<InvalidOperationException>(() => service.CreateSavedCase(kernelWorkbook, plan));
                Excel.Workbook openedWorkbook = Assert.Single(kernelWorkbook.Application.Workbooks);

                Assert.Equal(1, openedWorkbook.CloseCallCount);
                Assert.True(kernelWorkbook.Application.DisplayAlerts);
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
