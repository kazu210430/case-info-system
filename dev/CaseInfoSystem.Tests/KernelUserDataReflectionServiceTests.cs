using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    [Collection("ExcelApplicationCreatedApplications")]
    public class KernelUserDataReflectionServiceTests
    {
        private static readonly object IsolatedApplicationLock = new object();

        [Fact]
        public void ReflectToAccountingSetOnly_WhenTemplateIsNotOpen_UsesIsolatedLifecycleAndRestoresSharedUiState()
        {
            lock (IsolatedApplicationLock)
            {
                string tempDirectory = CreateTempDirectory();
                try
                {
                    Excel.Application.ResetCreatedApplications();

                    string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                    File.WriteAllText(templatePath, "template");

                    Excel.Application kernelApplication = new Excel.Application
                    {
                        DisplayAlerts = true,
                        EnableEvents = true,
                        ScreenUpdating = true,
                        StatusBar = "ready"
                    };
                    Excel.Workbook kernelWorkbook = CreateKernelWorkbook(tempDirectory, kernelApplication);
                    Excel.Worksheet userDataWorksheet = CreateUserDataWorksheet();
                    kernelWorkbook.Worksheets.Add(userDataWorksheet);

                    Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                    Excel.Application.ConfigureNewApplication = application =>
                        application.Workbooks.OpenBehavior = (_, __, ___) => accountingWorkbook;

                    var excelInteropService = new ExcelInteropService
                    {
                        OnFindOpenWorkbook = _ => null,
                        OnFindWorksheetByCodeName = (_, __) => userDataWorksheet,
                        OnReadKeyValueMapFromColumnsAandB = _ => CreateUserDataValues()
                    };
                    var service = CreateService(excelInteropService, templatePath, new AccountingWorkbookService(), new List<string>());

                    service.ReflectToAccountingSetOnly(CreateContext(kernelWorkbook, tempDirectory));

                    Excel.Application isolatedApplication = Excel.Application.CreatedApplications.Last();
                    var properties = Assert.IsAssignableFrom<IDictionary<string, string>>(accountingWorkbook.CustomDocumentProperties);

                    Assert.Equal(2, Excel.Application.CreatedApplications.Count);
                    Assert.NotSame(kernelApplication, isolatedApplication);
                    Assert.False(isolatedApplication.Visible);
                    Assert.False(isolatedApplication.DisplayAlerts);
                    Assert.False(isolatedApplication.ScreenUpdating);
                    Assert.False(isolatedApplication.EnableEvents);
                    Assert.Equal(1, accountingWorkbook.SaveCallCount);
                    Assert.Equal(1, accountingWorkbook.CloseCallCount);
                    Assert.Equal(1, isolatedApplication.QuitCallCount);
                    Assert.Equal(kernelWorkbook.FullName, properties[AccountingSetSpec.SourceKernelPathPropertyName]);
                    Assert.True(kernelApplication.DisplayAlerts);
                    Assert.True(kernelApplication.EnableEvents);
                    Assert.True(kernelApplication.ScreenUpdating);
                    Assert.Equal("ready", kernelApplication.StatusBar);
                }
                finally
                {
                    Excel.Application.ResetCreatedApplications();
                    TryDeleteDirectory(tempDirectory);
                }
            }
        }

        [Fact]
        public void ReflectToAccountingSetOnly_WhenIsolatedWriteFails_StillCleansUpAndRestoresSharedUiState()
        {
            lock (IsolatedApplicationLock)
            {
                string tempDirectory = CreateTempDirectory();
                try
                {
                    Excel.Application.ResetCreatedApplications();

                    string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                    File.WriteAllText(templatePath, "template");

                    Excel.Application kernelApplication = new Excel.Application
                    {
                        DisplayAlerts = true,
                        EnableEvents = true,
                        ScreenUpdating = true,
                        StatusBar = "ready"
                    };
                    Excel.Workbook kernelWorkbook = CreateKernelWorkbook(tempDirectory, kernelApplication);
                    Excel.Worksheet userDataWorksheet = CreateUserDataWorksheet();
                    kernelWorkbook.Worksheets.Add(userDataWorksheet);

                    Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                    Excel.Application.ConfigureNewApplication = application =>
                        application.Workbooks.OpenBehavior = (_, __, ___) => accountingWorkbook;

                    InvalidOperationException expected = new InvalidOperationException("write failed");
                    var excelInteropService = new ExcelInteropService
                    {
                        OnFindOpenWorkbook = _ => null,
                        OnFindWorksheetByCodeName = (_, __) => userDataWorksheet,
                        OnReadKeyValueMapFromColumnsAandB = _ => CreateUserDataValues()
                    };
                    var workbookService = new AccountingWorkbookService
                    {
                        OnWriteCell = (_, __, ___, ____) => throw expected
                    };
                    var service = CreateService(excelInteropService, templatePath, workbookService, new List<string>());

                    InvalidOperationException actual = Assert.Throws<InvalidOperationException>(
                        () => service.ReflectToAccountingSetOnly(CreateContext(kernelWorkbook, tempDirectory)));

                    Excel.Application isolatedApplication = Excel.Application.CreatedApplications.Last();

                    Assert.Same(expected, actual);
                    Assert.Equal(2, Excel.Application.CreatedApplications.Count);
                    Assert.NotSame(kernelApplication, isolatedApplication);
                    Assert.Equal(0, accountingWorkbook.SaveCallCount);
                    Assert.Equal(1, accountingWorkbook.CloseCallCount);
                    Assert.Equal(1, isolatedApplication.QuitCallCount);
                    Assert.True(kernelApplication.DisplayAlerts);
                    Assert.True(kernelApplication.EnableEvents);
                    Assert.True(kernelApplication.ScreenUpdating);
                    Assert.Equal("ready", kernelApplication.StatusBar);
                }
                finally
                {
                    Excel.Application.ResetCreatedApplications();
                    TryDeleteDirectory(tempDirectory);
                }
            }
        }

        private static KernelUserDataReflectionService CreateService(
            ExcelInteropService excelInteropService,
            string templatePath,
            AccountingWorkbookService accountingWorkbookService,
            List<string> logs)
        {
            var pathCompatibilityService = new PathCompatibilityService();
            return new KernelUserDataReflectionService(
                new KernelWorkbookService(
                    OrchestrationTestSupport.CreateKernelCaseInteractionState(logs),
                    OrchestrationTestSupport.CreateLogger(logs),
                    new KernelWorkbookService.KernelWorkbookServiceTestHooks()),
                excelInteropService,
                new AccountingTemplateResolver
                {
                    OnResolveTemplatePath = _ => templatePath
                },
                accountingWorkbookService,
                pathCompatibilityService,
                new UserDataBaseMappingRepository(excelInteropService),
                OrchestrationTestSupport.CreateLogger(logs));
        }

        private static WorkbookContext CreateContext(Excel.Workbook kernelWorkbook, string systemRoot)
        {
            return new WorkbookContext(
                kernelWorkbook,
                null,
                WorkbookRole.Kernel,
                systemRoot,
                kernelWorkbook.FullName,
                string.Empty);
        }

        private static Excel.Workbook CreateKernelWorkbook(string tempDirectory, Excel.Application application)
        {
            var workbook = new Excel.Workbook
            {
                Application = application,
                FullName = Path.Combine(tempDirectory, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsx")),
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsx"),
                Path = tempDirectory,
                CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = tempDirectory
                }
            };
            return workbook;
        }

        private static Excel.Workbook CreateWorkbook(string fullPath)
        {
            return new Excel.Workbook
            {
                FullName = fullPath,
                Name = Path.GetFileName(fullPath),
                Path = Path.GetDirectoryName(fullPath) ?? string.Empty
            };
        }

        private static Excel.Worksheet CreateUserDataWorksheet()
        {
            return new Excel.Worksheet
            {
                CodeName = AccountingSetSpec.UserDataSheetCodeName,
                Name = AccountingSetSpec.UserDataSheetName
            };
        }

        private static IReadOnlyDictionary<string, string> CreateUserDataValues()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["当方_郵便番号"] = "100-0001",
                ["当方_住所"] = "東京都千代田区1-1",
                ["当方_事務所名"] = "OpenAI法律事務所",
                ["当方_電話"] = "03-1111-2222",
                ["銀行・支店"] = "みずほ銀行 東京支店",
                ["口座番号・名義"] = "1234567 OpenAI"
            };
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
    }
}
