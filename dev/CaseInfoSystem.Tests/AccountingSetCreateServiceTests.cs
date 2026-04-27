using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class AccountingSetCreateServiceTests
    {
        private const string CustomerNameKey = "顧客_名前";
        private const string CustomerHonorificKey = "顧客_敬称";
        private const string LawyerKey = "当方_弁護士";

        [Fact]
        public void Execute_WhenInputsAreValid_CopiesTemplateAndWritesInitialWorkbookData()
        {
            List<string> logs = new List<string>();
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "template.xlsx");
                string outputPath = Path.Combine(tempDirectory, "accounting-set.xlsx");
                File.WriteAllText(templatePath, "template");

                Excel.Workbook caseWorkbook = CreateWorkbook(
                    Path.Combine(tempDirectory, "case.xlsx"),
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["SYSTEM_ROOT"] = tempDirectory,
                        ["NAME_RULE_A"] = "2026",
                        ["NAME_RULE_B"] = "ACC"
                    });
                Excel.Workbook createdWorkbook = CreateWorkbook(outputPath);

                var caseContextFactory = new CaseContextFactory
                {
                    OnCreateForDocumentCreate = _ => new CaseContext
                    {
                        CaseValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            [CustomerNameKey] = "Alpha",
                            [CustomerHonorificKey] = "御中",
                            [LawyerKey] = "弁護士A\r\n弁護士B"
                        }
                    }
                };
                var documentOutputService = new DocumentOutputService(new ExcelInteropService(), new PathCompatibilityService(), OrchestrationTestSupport.CreateLogger(logs))
                {
                    OnResolveWorkbookFolder = _ => tempDirectory
                };
                string capturedFolder = null;
                string capturedCustomerName = null;
                string capturedTemplate = null;
                var accountingSetNamingService = new AccountingSetNamingService
                {
                    OnBuildCaseOutputPath = (_, folder, customerName, template) =>
                    {
                        capturedFolder = folder;
                        capturedCustomerName = customerName;
                        capturedTemplate = template;
                        return outputPath;
                    }
                };
                var accountingTemplateResolver = new AccountingTemplateResolver
                {
                    OnResolveTemplatePath = _ => templatePath
                };
                List<(string Sheet, string Address, string Value)> cellWrites = new List<(string Sheet, string Address, string Value)>();
                List<(IReadOnlyList<string> Sheets, string Address, string Value)> multiWrites = new List<(IReadOnlyList<string> Sheets, string Address, string Value)>();
                string reflectedLawyers = null;
                bool activatedInvoiceEntry = false;
                var accountingWorkbookService = new AccountingWorkbookService
                {
                    OnOpenInCurrentApplication = path =>
                    {
                        Assert.Equal(outputPath, path);
                        return createdWorkbook;
                    },
                    OnWriteSameValueToSheets = (_, sheets, address, value) => multiWrites.Add((sheets.ToArray(), address, value)),
                    OnWriteCell = (_, sheet, address, value) => cellWrites.Add((sheet, address, value)),
                    OnReflectLawyers = (_, lawyerLines) =>
                    {
                        reflectedLawyers = lawyerLines;
                        return new AccountingLawyerMappingResult();
                    },
                    OnActivateInvoiceEntry = _ => activatedInvoiceEntry = true
                };
                var transientPaneSuppressionService = CreateTransientPaneSuppressionService(outputPath, logs);
                AccountingSetPresentationWaitService.WaitSession waitSession = new AccountingSetPresentationWaitService.WaitSession();
                var waitService = new AccountingSetPresentationWaitService
                {
                    OnShowWaiting = _ => waitSession
                };
                Excel.Workbook shownWorkbook = null;
                Globals.ThisAddIn = new ThisAddIn
                {
                    ShowWorkbookTaskPaneWhenReadyHandler = (workbook, _) => shownWorkbook = workbook
                };

                var service = new AccountingSetCreateService(
                    new ExcelInteropService(),
                    caseContextFactory,
                    documentOutputService,
                    accountingSetNamingService,
                    accountingTemplateResolver,
                    accountingWorkbookService,
                    new PathCompatibilityService(),
                    transientPaneSuppressionService,
                    waitService,
                    OrchestrationTestSupport.CreateLogger(logs));

                service.Execute(caseWorkbook);

                Assert.True(File.Exists(outputPath));
                Assert.Equal(tempDirectory, capturedFolder);
                Assert.Equal("Alpha", capturedCustomerName);
                Assert.Equal(templatePath, capturedTemplate);
                Assert.Equal("弁護士A\r\n弁護士B", reflectedLawyers);
                Assert.True(activatedInvoiceEntry);
                Assert.Same(createdWorkbook, shownWorkbook);
                Assert.False(transientPaneSuppressionService.IsSuppressedPath(outputPath));
                Assert.Contains(AccountingSetPresentationWaitService.CreatingStageTitle, waitSession.Stages);
                Assert.Contains(AccountingSetPresentationWaitService.OpeningWorkbookStageTitle, waitSession.Stages);
                Assert.Contains(AccountingSetPresentationWaitService.ApplyingInitialDataStageTitle, waitSession.Stages);
                Assert.Contains(AccountingSetPresentationWaitService.ShowingInputScreenStageTitle, waitSession.Stages);

                var properties = Assert.IsAssignableFrom<IDictionary<string, string>>(createdWorkbook.CustomDocumentProperties);
                Assert.Equal(AccountingSetSpec.WorkbookKindAccountingSetValue, properties[AccountingSetSpec.WorkbookKindPropertyName]);
                Assert.Equal(caseWorkbook.FullName, properties[AccountingSetSpec.SourceCasePathPropertyName]);
                Assert.Equal(tempDirectory, properties["SYSTEM_ROOT"]);
                Assert.Equal("2026", properties["NAME_RULE_A"]);
                Assert.Equal("ACC", properties["NAME_RULE_B"]);

                Assert.Contains(multiWrites, write =>
                    write.Address == AccountingSetSpec.CustomerWriteCellAddress
                    && write.Value == "Alpha 御中"
                    && write.Sheets.SequenceEqual(new[]
                    {
                        AccountingSetSpec.EstimateSheetName,
                        AccountingSetSpec.InvoiceSheetName,
                        AccountingSetSpec.ReceiptSheetName
                    }));
                Assert.Contains(cellWrites, write =>
                    write.Sheet == AccountingSetSpec.AccountingRequestSheetName
                    && write.Address == AccountingSetSpec.CustomerWriteCellAddress
                    && write.Value == "Alpha 御中");
            }
            finally
            {
                Globals.ThisAddIn = new ThisAddIn();
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenCaseDataIsEmpty_ThrowsInvalidOperationException()
        {
            var service = CreateService(
                caseValues: new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                templatePath: "ignored",
                outputPath: "ignored",
                outputFolderPath: "ignored");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => service.Service.Execute(service.CaseWorkbook));

            Assert.Contains("案件データ", exception.Message);
        }

        [Fact]
        public void Execute_WhenTemplateFileIsMissing_PropagatesFileNotFoundException()
        {
            string tempDirectory = CreateTempDirectory();
            try
            {
                string missingTemplatePath = Path.Combine(tempDirectory, "missing.xlsx");
                string outputPath = Path.Combine(tempDirectory, "accounting-set.xlsx");
                var service = CreateService(
                    caseValues: CreateCaseValues(),
                    templatePath: missingTemplatePath,
                    outputPath: outputPath,
                    outputFolderPath: tempDirectory);

                Assert.Throws<FileNotFoundException>(() => service.Service.Execute(service.CaseWorkbook));
                Assert.False(File.Exists(outputPath));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
                Globals.ThisAddIn = new ThisAddIn();
            }
        }

        [Fact]
        public void Execute_WhenOutputAlreadyExists_DeletesExistingFileDuringFailureCleanup()
        {
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "template.xlsx");
                string outputPath = Path.Combine(tempDirectory, "accounting-set.xlsx");
                File.WriteAllText(templatePath, "template");
                File.WriteAllText(outputPath, "existing");
                var service = CreateService(
                    caseValues: CreateCaseValues(),
                    templatePath: templatePath,
                    outputPath: outputPath,
                    outputFolderPath: tempDirectory);

                Assert.Throws<IOException>(() => service.Service.Execute(service.CaseWorkbook));
                Assert.False(File.Exists(outputPath));
                Assert.False(service.TransientPaneSuppressionService.IsSuppressedPath(outputPath));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
                Globals.ThisAddIn = new ThisAddIn();
            }
        }

        [Fact]
        public void Execute_WhenOpenWorkbookFails_RethrowsOriginalExceptionAndDeletesCopiedOutput()
        {
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "template.xlsx");
                string outputPath = Path.Combine(tempDirectory, "accounting-set.xlsx");
                File.WriteAllText(templatePath, "template");
                InvalidOperationException expected = new InvalidOperationException("open failed");
                var service = CreateService(
                    caseValues: CreateCaseValues(),
                    templatePath: templatePath,
                    outputPath: outputPath,
                    outputFolderPath: tempDirectory,
                    openWorkbook: _ => throw expected);

                InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() => service.Service.Execute(service.CaseWorkbook));

                Assert.Same(expected, actual);
                Assert.False(File.Exists(outputPath));
                Assert.False(service.TransientPaneSuppressionService.IsSuppressedPath(outputPath));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
                Globals.ThisAddIn = new ThisAddIn();
            }
        }

        private static IDictionary<string, string> CreateCaseValues()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [CustomerNameKey] = "Alpha",
                [CustomerHonorificKey] = "御中",
                [LawyerKey] = "弁護士A"
            };
        }

        private static TestServiceContext CreateService(
            IDictionary<string, string> caseValues,
            string templatePath,
            string outputPath,
            string outputFolderPath,
            Func<string, Excel.Workbook> openWorkbook = null)
        {
            List<string> logs = new List<string>();
            var caseWorkbook = CreateWorkbook(
                Path.Combine(outputFolderPath, "case.xlsx"),
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = outputFolderPath,
                    ["NAME_RULE_A"] = "2026",
                    ["NAME_RULE_B"] = "ACC"
                });
            var caseContextFactory = new CaseContextFactory
            {
                OnCreateForDocumentCreate = _ => new CaseContext
                {
                    CaseValues = caseValues == null
                        ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        : new Dictionary<string, string>(caseValues, StringComparer.OrdinalIgnoreCase)
                }
            };
            var documentOutputService = new DocumentOutputService(new ExcelInteropService(), new PathCompatibilityService(), OrchestrationTestSupport.CreateLogger(logs))
            {
                OnResolveWorkbookFolder = _ => outputFolderPath
            };
            var namingService = new AccountingSetNamingService
            {
                OnBuildCaseOutputPath = (_, __, ___, ____) => outputPath
            };
            var templateResolver = new AccountingTemplateResolver
            {
                OnResolveTemplatePath = _ => templatePath
            };
            var workbookService = new AccountingWorkbookService
            {
                OnOpenInCurrentApplication = openWorkbook ?? (_ => CreateWorkbook(outputPath)),
                OnReflectLawyers = (_, __) => new AccountingLawyerMappingResult()
            };
            var transientPaneSuppressionService = CreateTransientPaneSuppressionService(outputPath, logs);
            var waitService = new AccountingSetPresentationWaitService();
            Globals.ThisAddIn = new ThisAddIn();

            return new TestServiceContext
            {
                CaseWorkbook = caseWorkbook,
                Service = new AccountingSetCreateService(
                    new ExcelInteropService(),
                    caseContextFactory,
                    documentOutputService,
                    namingService,
                    templateResolver,
                    workbookService,
                    new PathCompatibilityService(),
                    transientPaneSuppressionService,
                    waitService,
                    OrchestrationTestSupport.CreateLogger(logs)),
                TransientPaneSuppressionService = transientPaneSuppressionService
            };
        }

        private static Excel.Workbook CreateWorkbook(string fullPath, IDictionary<string, string> properties = null)
        {
            return new Excel.Workbook
            {
                FullName = fullPath,
                Name = Path.GetFileName(fullPath),
                Path = Path.GetDirectoryName(fullPath) ?? string.Empty,
                CustomDocumentProperties = properties
            };
        }

        private static TransientPaneSuppressionService CreateTransientPaneSuppressionService(string workbookFullName, List<string> logs)
        {
            var excelInteropService = new FakeExcelInteropService
            {
                WorkbookFullName = workbookFullName
            };
            return new TransientPaneSuppressionService(excelInteropService, new PathCompatibilityService(), OrchestrationTestSupport.CreateLogger(logs));
        }

        private static string CreateTempDirectory()
        {
            string path = Path.Combine(Path.GetTempPath(), "CaseInfoSystem.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                {
                    Directory.Delete(path, recursive: true);
                }
            }
            catch
            {
            }
        }

        private sealed class TestServiceContext
        {
            internal Excel.Workbook CaseWorkbook { get; set; }

            internal AccountingSetCreateService Service { get; set; }

            internal TransientPaneSuppressionService TransientPaneSuppressionService { get; set; }
        }
    }
}
