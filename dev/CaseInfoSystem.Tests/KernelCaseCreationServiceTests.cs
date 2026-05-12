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
        public void CreateCase_WhenBoundKernelWorkbookIsProvided_UsesItWithoutOpenKernelFallback()
        {
            List<string> logs = new List<string>();
            string tempRoot = Path.Combine(Path.GetTempPath(), "KernelCaseCreationServiceTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            string caseFolderPath = Path.Combine(tempRoot, "Selected");
            Directory.CreateDirectory(caseFolderPath);
            string baseWorkbookPath = Path.Combine(tempRoot, "Base.xlsx");
            File.WriteAllText(baseWorkbookPath, "base");

            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                FullName = Path.Combine(tempRoot, "Kernel.xlsm"),
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                Path = tempRoot,
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["NAME_RULE_A"] = "YYYY",
                    ["NAME_RULE_B"] = "DOC"
                }
            };
            kernelWorkbook.Application = new Excel.Application
            {
                DisplayAlerts = true
            };

            int openKernelCalls = 0;
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                });

            Excel.Application hiddenApplication = new Excel.Application();
            Excel.Workbook hiddenWorkbook = new Excel.Workbook
            {
                Application = hiddenApplication
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
            var caseWorkbookOpenStrategy = CreateCaseWorkbookOpenStrategy(hiddenApplication, hiddenWorkbook, logs);
            var casePathService = new KernelCasePathService
            {
                OnResolveSystemRoot = workbook => workbook == null ? string.Empty : workbook.Path,
                OnResolveBaseWorkbookPath = _ => baseWorkbookPath,
                OnResolveCaseFolderPath = (_, __) => caseFolderPath,
                OnEnsureFolderExists = folderPath =>
                {
                    Directory.CreateDirectory(folderPath);
                    return true;
                },
                OnResolveCaseWorkbookExtension = _ => ".xlsx",
                OnBuildCaseWorkbookPath = (folderPath, caseWorkbookName) => Path.Combine(folderPath, caseWorkbookName)
            };

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                casePathService,
                new CaseWorkbookInitializer(),
                caseWorkbookOpenStrategy,
                transientPaneSuppressionService,
                caseWorkbookLifecycleService,
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(logs));

            try
            {
                KernelCaseCreationResult result = service.CreateCase(
                    kernelWorkbook,
                    tempRoot,
                    new KernelCaseCreationRequest
                    {
                        CustomerName = "ClientA",
                        Mode = KernelCaseCreationMode.CreateCaseBatch,
                        SelectedFolderPath = caseFolderPath
                    });

                Assert.True(result.Success);
                Assert.Equal(0, openKernelCalls);
                Assert.Equal(caseFolderPath, result.CaseFolderPath);
                Assert.Equal(caseFolderPath, Path.GetDirectoryName(result.CaseWorkbookPath));
                Assert.True(File.Exists(result.CaseWorkbookPath));
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
        public void CreateCase_WhenBoundKernelRootMismatches_FailsClosedWithoutOpenKernelFallback()
        {
            int openKernelCalls = 0;
            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\actual-root\Kernel.xlsm",
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                Path = @"C:\actual-root"
            };
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                });
            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService
                {
                    OnResolveSystemRoot = workbook => workbook == null ? string.Empty : workbook.Path
                },
                new CaseWorkbookInitializer(),
                CreateCaseWorkbookOpenStrategy(new Excel.Application(), new Excel.Workbook(), new List<string>()),
                new TransientPaneSuppressionService(new FakeExcelInteropService(), new PathCompatibilityService(), OrchestrationTestSupport.CreateLogger(new List<string>())),
                new CaseWorkbookLifecycleService(
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks()),
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()));

            Assert.Throws<InvalidOperationException>(() => service.CreateCase(
                kernelWorkbook,
                @"C:\expected-root",
                new KernelCaseCreationRequest
                {
                    CustomerName = "ClientA",
                    Mode = KernelCaseCreationMode.CreateCaseBatch,
                    SelectedFolderPath = @"C:\cases"
                }));
            Assert.Equal(0, openKernelCalls);
        }

        [Theory]
        [InlineData(Excel.XlWindowState.xlMinimized)]
        [InlineData(Excel.XlWindowState.xlNormal)]
        public void CreateSavedCase_ForInteractiveModes_DoesNotMakeHiddenSessionWindowVisibleBeforeSave(Excel.XlWindowState initialWindowState)
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
                FullName = @"C:\System\譯井ｻｶ諠・ｱSystem.xlsm",
                Name = "譯井ｻｶ諠・ｱSystem.xlsm",
                Path = @"C:\System"
            };
            kernelWorkbook.Application = new Excel.Application();
            var kernelWorkbookService = new KernelWorkbookService(
                OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                new KernelWorkbookService.KernelWorkbookServiceTestHooks
                {
                });

            var caseWorkbookInitializer = new CaseWorkbookInitializer
            {
                OnInitializeForVisibleCreate = (_, createdCaseWorkbook, _) =>
                {
                    initializedPaths.Add(createdCaseWorkbook == null ? string.Empty : createdCaseWorkbook.FullName ?? string.Empty);
                },
                OnInitializeForHiddenCreate = (_, _, _) => throw new InvalidOperationException("interactive flow must preserve visible-create startup state")
            };

            Excel.Application hiddenApplication = new Excel.Application();
            Excel.Workbook hiddenWorkbook = new Excel.Workbook
            {
                FullName = finalCaseWorkbookPath,
                Name = Path.GetFileName(finalCaseWorkbookPath),
                Path = caseFolderPath,
                Application = hiddenApplication
            };
            Excel.Worksheet homeWorksheet = new Excel.Worksheet
            {
                CodeName = "shHOME",
                Name = "\u30db\u30fc\u30e0",
                Parent = hiddenWorkbook
            };
            Excel.Window hiddenWindow = new Excel.Window
            {
                Visible = false,
                WindowState = initialWindowState
            };
            hiddenWorkbook.Worksheets.Add(homeWorksheet);
            hiddenWorkbook.ActiveSheet = homeWorksheet;
            hiddenWorkbook.Windows.Add(hiddenWindow);
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

            var caseWorkbookOpenStrategy = CreateCaseWorkbookOpenStrategy(hiddenApplication, hiddenWorkbook, logs);

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                caseWorkbookInitializer,
                caseWorkbookOpenStrategy,
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

                Assert.Equal(new[] { finalCaseWorkbookPath }, initializedPaths);
                Assert.Empty(kernelWorkbook.Application.Workbooks);
                Assert.Equal(1, hiddenWorkbook.SaveCallCount);
                Assert.Equal(1, hiddenWorkbook.CloseCallCount);
                Assert.False(hiddenWindow.Visible);
                Assert.Equal(Excel.XlWindowState.xlNormal, hiddenWindow.WindowState);
                Assert.False(hiddenWindow.Activated);
                Assert.Equal(1, hiddenApplication.QuitCallCount);
                Assert.True(result.Success);
                Assert.Equal(finalCaseWorkbookPath, result.CaseWorkbookPath);
                Assert.True(File.Exists(finalCaseWorkbookPath));
                Assert.Contains(logs, message => message.IndexOf("hidden session opened", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.Contains(logs, message => message.IndexOf("deferred visible presentation", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.Contains(logs, message => message.IndexOf("save-window-visible-deferred", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.Contains(logs, message => message.IndexOf("interactive CASE workbook window normalized before save", StringComparison.OrdinalIgnoreCase) >= 0);
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
                });

            Excel.Application hiddenApplication = new Excel.Application();
            Excel.Workbook hiddenWorkbook = new Excel.Workbook
            {
                FullName = caseWorkbookPath,
                Name = Path.GetFileName(caseWorkbookPath),
                Path = tempRoot,
                Application = hiddenApplication
            };
            Excel.Worksheet homeWorksheet = new Excel.Worksheet
            {
                CodeName = "shHOME",
                Name = "\u30db\u30fc\u30e0",
                Parent = hiddenWorkbook
            };
            hiddenWorkbook.Worksheets.Add(homeWorksheet);
            hiddenWorkbook.ActiveSheet = homeWorksheet;
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

            var caseWorkbookOpenStrategy = CreateCaseWorkbookOpenStrategy(hiddenApplication, hiddenWorkbook, logs);

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                new CaseWorkbookInitializer(),
                caseWorkbookOpenStrategy,
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

                Assert.True(result.Success);
                Assert.True(kernelWorkbook.Application.DisplayAlerts);
                Assert.Equal(1, hiddenWorkbook.CloseCallCount);
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
                });

            var caseWorkbookInitializer = new CaseWorkbookInitializer
            {
                OnInitializeForVisibleCreate = (_, _, _) => throw new InvalidOperationException("boom"),
                OnInitializeForHiddenCreate = (_, _, _) => throw new InvalidOperationException("boom")
            };

            Excel.Application hiddenApplication = new Excel.Application();
            Excel.Workbook hiddenWorkbook = new Excel.Workbook
            {
                FullName = caseWorkbookPath,
                Name = Path.GetFileName(caseWorkbookPath),
                Path = tempRoot,
                Application = hiddenApplication
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

            var caseWorkbookOpenStrategy = CreateCaseWorkbookOpenStrategy(hiddenApplication, hiddenWorkbook, logs);

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                caseWorkbookInitializer,
                caseWorkbookOpenStrategy,
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

                Assert.Equal(1, hiddenWorkbook.CloseCallCount);
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

        [Fact]
        public void CreateSavedCase_ForBatch_UsesHiddenSessionAndHiddenInitializer()
        {
            List<string> logs = new List<string>();
            List<string> initializedPaths = new List<string>();
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
                });

            int visibleInitializeCount = 0;
            var caseWorkbookInitializer = new CaseWorkbookInitializer
            {
                OnInitializeForVisibleCreate = (_, _, _) => visibleInitializeCount++,
                OnInitializeForHiddenCreate = (_, createdCaseWorkbook, _) =>
                {
                    initializedPaths.Add(createdCaseWorkbook == null ? string.Empty : createdCaseWorkbook.FullName ?? string.Empty);
                }
            };

            Excel.Application hiddenApplication = new Excel.Application();
            Excel.Workbook hiddenWorkbook = new Excel.Workbook
            {
                FullName = caseWorkbookPath,
                Name = Path.GetFileName(caseWorkbookPath),
                Path = tempRoot,
                Application = hiddenApplication
            };
            Excel.Worksheet homeWorksheet = new Excel.Worksheet
            {
                CodeName = "shHOME",
                Name = "\u30db\u30fc\u30e0",
                Parent = hiddenWorkbook
            };
            Excel.Window hiddenWindow = new Excel.Window
            {
                Visible = false,
                WindowState = Excel.XlWindowState.xlMinimized
            };
            hiddenWorkbook.Worksheets.Add(homeWorksheet);
            hiddenWorkbook.ActiveSheet = homeWorksheet;
            hiddenWorkbook.Windows.Add(hiddenWindow);
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

            var caseWorkbookOpenStrategy = CreateCaseWorkbookOpenStrategy(hiddenApplication, hiddenWorkbook, logs);

            var service = new KernelCaseCreationService(
                kernelWorkbookService,
                new KernelCasePathService(),
                caseWorkbookInitializer,
                caseWorkbookOpenStrategy,
                transientPaneSuppressionService,
                caseWorkbookLifecycleService,
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(logs));

            var plan = new KernelCaseCreationPlan
            {
                Mode = KernelCaseCreationMode.CreateCaseBatch,
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

                Assert.True(result.Success);
                Assert.Equal(0, visibleInitializeCount);
                Assert.Equal(new[] { caseWorkbookPath }, initializedPaths);
                Assert.Equal(1, hiddenWorkbook.SaveCallCount);
                Assert.Equal(1, hiddenWorkbook.CloseCallCount);
                Assert.Equal(1, hiddenWorkbook.Windows.Count);
                Assert.True(hiddenWindow.Visible);
                Assert.Equal(Excel.XlWindowState.xlNormal, hiddenWindow.WindowState);
                Assert.False(hiddenWindow.Activated);
                Assert.Equal(1, hiddenApplication.QuitCallCount);
                Assert.True(kernelWorkbook.Application.DisplayAlerts);
                Assert.Contains(logs, message => message.IndexOf("hidden session opened", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            finally
            {
                if (Directory.Exists(tempRoot))
                {
                    Directory.Delete(tempRoot, recursive: true);
                }
            }
        }

        private static CaseWorkbookOpenStrategy CreateCaseWorkbookOpenStrategy(Excel.Application hiddenApplication, Excel.Workbook hiddenWorkbook, List<string> logs)
        {
            hiddenApplication.Workbooks.OpenBehavior = (filename, _, _) =>
            {
                hiddenWorkbook.FullName = filename ?? string.Empty;
                hiddenWorkbook.Name = Path.GetFileName(filename ?? string.Empty);
                hiddenWorkbook.Path = Path.GetDirectoryName(filename ?? string.Empty) ?? string.Empty;
                hiddenWorkbook.Application = hiddenApplication;
                return hiddenWorkbook;
            };

            return new CaseWorkbookOpenStrategy(
                new Excel.Application(),
                new WorkbookRoleResolver(),
                OrchestrationTestSupport.CreateLogger(logs),
                hiddenApplicationFactory: () => hiddenApplication);
        }
    }
}
