using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class AccountingSetKernelSyncServiceTests
    {
        [Fact]
        public void Execute_WhenUserDataExists_WritesMappedPropertiesAndCells()
        {
            List<string> logs = new List<string>();
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                Excel.Worksheet userDataWorksheet = CreateUserDataWorksheet();
                kernelWorkbook.Worksheets.Add(userDataWorksheet);
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);

                var excelInteropService = new ExcelInteropService
                {
                    OnFindOpenWorkbook = path => string.Equals(path, templatePath, StringComparison.OrdinalIgnoreCase) ? accountingWorkbook : null
                };
                var templateResolver = new AccountingTemplateResolver
                {
                    OnResolveTemplatePath = _ => templatePath
                };
                List<(IReadOnlyList<string> Sheets, string Address, string Value)> rangeWrites = new List<(IReadOnlyList<string> Sheets, string Address, string Value)>();
                List<(string Sheet, string Address, string Value)> cellWrites = new List<(string Sheet, string Address, string Value)>();
                var workbookService = new AccountingWorkbookService
                {
                    OnWriteSameValueToSheets = (_, sheets, address, value) => rangeWrites.Add((sheets.ToArray(), address, value)),
                    OnWriteCell = (_, sheet, address, value) => cellWrites.Add((sheet, address, value))
                };
                var service = new AccountingSetKernelSyncService(
                    excelInteropService,
                    templateResolver,
                    workbookService,
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(logs));

                service.Execute(kernelWorkbook);

                var properties = Assert.IsAssignableFrom<IDictionary<string, string>>(accountingWorkbook.CustomDocumentProperties);
                Assert.Equal(kernelWorkbook.FullName, properties[AccountingSetSpec.SourceKernelPathPropertyName]);
                Assert.Equal(1, accountingWorkbook.SaveCallCount);
                Assert.Equal(0, accountingWorkbook.CloseCallCount);
                Assert.Contains(rangeWrites, write =>
                    write.Address == AccountingSetSpec.AccountingAddressCellAddress
                    && write.Sheets.SequenceEqual(new[]
                    {
                        AccountingSetSpec.EstimateSheetName,
                        AccountingSetSpec.InvoiceSheetName,
                        AccountingSetSpec.ReceiptSheetName
                    })
                    && write.Value.Contains("0")
                    && write.Value.Contains("東京都")
                    && write.Value.Contains("03-1111-2222"));
                Assert.Contains(rangeWrites, write =>
                    write.Address == AccountingSetSpec.InstallmentAddressCellAddress
                    && write.Sheets.SequenceEqual(new[]
                    {
                        AccountingSetSpec.InstallmentSheetName,
                        AccountingSetSpec.PaymentHistorySheetName
                    })
                    && write.Value.Contains("\n"));
                Assert.Contains(cellWrites, write =>
                    write.Sheet == AccountingSetSpec.InvoiceSheetName
                    && write.Address == AccountingSetSpec.InvoiceNameRow1CellAddress
                    && write.Value == "銀行支店ABC");
                Assert.Contains(cellWrites, write =>
                    write.Sheet == AccountingSetSpec.InvoiceSheetName
                    && write.Address == AccountingSetSpec.InvoiceNameRow2CellAddress
                    && write.Value == "名義XYZ");
                Assert.Contains(cellWrites, write =>
                    write.Sheet == AccountingSetSpec.InstallmentSheetName
                    && write.Address == AccountingSetSpec.InstallmentNameRow1CellAddress
                    && write.Value == "銀行支店ABC");
                Assert.Contains(cellWrites, write =>
                    write.Sheet == AccountingSetSpec.InstallmentSheetName
                    && write.Address == AccountingSetSpec.InstallmentAddressCellAddress
                    && write.Value.Contains("\n"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "東京都"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "千代田区1-1"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "03-1111-2222"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "銀行支店ABC"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "名義XYZ"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, tempDirectory));
                Assert.Contains(logs, message =>
                    ContainsOrdinal(message, "addressLinePresent=True")
                    && ContainsOrdinal(message, "addressLineLength=")
                    && ContainsOrdinal(message, "nameLine1Present=True")
                    && ContainsOrdinal(message, "propertyKey=SOURCE_KERNEL_PATH"));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenUserDataReadFails_TreatsMissingValueAsBlank()
        {
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                Excel.Worksheet userDataWorksheet = CreateUserDataWorksheet();
                userDataWorksheet.Cells.ThrowOnAccess(AccountingSetSpec.UserDataFirstDataRow, "B");
                kernelWorkbook.Worksheets.Add(userDataWorksheet);
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);

                var excelInteropService = new ExcelInteropService
                {
                    OnFindOpenWorkbook = _ => accountingWorkbook
                };
                var templateResolver = new AccountingTemplateResolver
                {
                    OnResolveTemplatePath = _ => templatePath
                };
                string addressLine = null;
                var workbookService = new AccountingWorkbookService
                {
                    OnWriteSameValueToSheets = (_, sheets, address, value) =>
                    {
                        if (address == AccountingSetSpec.AccountingAddressCellAddress
                            && sheets.Contains(AccountingSetSpec.InvoiceSheetName))
                        {
                            addressLine = value;
                        }
                    }
                };
                var service = new AccountingSetKernelSyncService(
                    excelInteropService,
                    templateResolver,
                    workbookService,
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(new List<string>()));

                service.Execute(kernelWorkbook);

                Assert.NotNull(addressLine);
                Assert.DoesNotContain("100-0001", addressLine);
                Assert.Contains("東京都", addressLine);
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenWriteFails_RethrowsOriginalException()
        {
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                kernelWorkbook.Worksheets.Add(CreateUserDataWorksheet());
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                InvalidOperationException expected = new InvalidOperationException("write failed");

                var service = new AccountingSetKernelSyncService(
                    new ExcelInteropService
                    {
                        OnFindOpenWorkbook = _ => accountingWorkbook
                    },
                    new AccountingTemplateResolver
                    {
                        OnResolveTemplatePath = _ => templatePath
                    },
                    new AccountingWorkbookService
                    {
                        OnWriteCell = (_, __, ___, ____) => throw expected
                    },
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(new List<string>()));

                InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() => service.Execute(kernelWorkbook));

                Assert.Same(expected, actual);
                Assert.Equal(0, accountingWorkbook.SaveCallCount);
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenTemplateIsNotOpen_OpensHiddenInCurrentApplicationAndRestoresState()
        {
            List<string> logs = new List<string>();
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Application sharedApplication = new Excel.Application
                {
                    DisplayAlerts = true,
                    ScreenUpdating = true,
                    EnableEvents = true
                };
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                kernelWorkbook.Application = sharedApplication;
                sharedApplication.ActiveWorkbook = kernelWorkbook;
                kernelWorkbook.Worksheets.Add(CreateUserDataWorksheet());
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                bool? displayAlertsDuringOpen = null;
                bool? screenUpdatingDuringOpen = null;
                bool? enableEventsDuringOpen = null;
                bool? hiddenVisibility = null;

                var service = new AccountingSetKernelSyncService(
                    new ExcelInteropService(),
                    new AccountingTemplateResolver
                    {
                        OnResolveTemplatePath = _ => templatePath
                    },
                    new AccountingWorkbookService
                    {
                        OnOpenInCurrentApplication = path =>
                        {
                            Assert.Equal(templatePath, path);
                            displayAlertsDuringOpen = sharedApplication.DisplayAlerts;
                            screenUpdatingDuringOpen = sharedApplication.ScreenUpdating;
                            enableEventsDuringOpen = sharedApplication.EnableEvents;
                            accountingWorkbook.Application = sharedApplication;
                            return accountingWorkbook;
                        },
                        OnSetWorkbookWindowsVisible = (_, visible) => hiddenVisibility = visible
                    },
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(logs));

                service.Execute(kernelWorkbook);

                Assert.False(displayAlertsDuringOpen.GetValueOrDefault(true));
                Assert.False(screenUpdatingDuringOpen.GetValueOrDefault(true));
                Assert.False(enableEventsDuringOpen.GetValueOrDefault(true));
                Assert.False(hiddenVisibility.GetValueOrDefault(true));
                Assert.True(sharedApplication.DisplayAlerts);
                Assert.True(sharedApplication.ScreenUpdating);
                Assert.True(sharedApplication.EnableEvents);
                Assert.Equal(1, accountingWorkbook.SaveCallCount);
                Assert.Equal(1, accountingWorkbook.CloseCallCount);
                Assert.Equal(0, sharedApplication.QuitCallCount);
                Assert.Contains(logs, message =>
                    ContainsOrdinal(message, "Accounting set kernel sync owned workbook close attempted.")
                    && ContainsOrdinal(message, "route=AccountingSetKernelSyncService.current-application-owned-workbook")
                    && ContainsOrdinal(message, "ownedWorkbook=True")
                    && ContainsOrdinal(message, "closeAttempted=True")
                    && ContainsOrdinal(message, "closeSucceeded=True")
                    && ContainsOrdinal(message, "hiddenWindowRoute=True"));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenCurrentApplicationFallbackCloseFails_LogsFailureOutcomeWithoutThrowing()
        {
            List<string> logs = new List<string>();
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Application sharedApplication = new Excel.Application
                {
                    DisplayAlerts = true,
                    ScreenUpdating = true,
                    EnableEvents = true
                };
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                kernelWorkbook.Application = sharedApplication;
                kernelWorkbook.Worksheets.Add(CreateUserDataWorksheet());
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                accountingWorkbook.CloseBehavior = () => throw new InvalidOperationException("close failed " + templatePath + " 東京都");

                var service = new AccountingSetKernelSyncService(
                    new ExcelInteropService(),
                    new AccountingTemplateResolver
                    {
                        OnResolveTemplatePath = _ => templatePath
                    },
                    new AccountingWorkbookService
                    {
                        OnOpenInCurrentApplication = _ =>
                        {
                            accountingWorkbook.Application = sharedApplication;
                            return accountingWorkbook;
                        }
                    },
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(logs));

                service.Execute(kernelWorkbook);

                Assert.Equal(1, accountingWorkbook.SaveCallCount);
                Assert.Equal(1, accountingWorkbook.CloseCallCount);
                Assert.Contains(logs, message =>
                    ContainsOrdinal(message, "WARN: Accounting set kernel sync owned workbook close attempted.")
                    && ContainsOrdinal(message, "route=AccountingSetKernelSyncService.current-application-owned-workbook")
                    && ContainsOrdinal(message, "ownedWorkbook=True")
                    && ContainsOrdinal(message, "closeAttempted=True")
                    && ContainsOrdinal(message, "closeSucceeded=False")
                    && ContainsOrdinal(message, "exceptionType=System.InvalidOperationException")
                    && ContainsOrdinal(message, "hresult=0x"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, tempDirectory));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "close failed"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "東京都"));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenCurrentApplicationFallbackWriteFails_StillClosesWorkbookAndRestoresState()
        {
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Application sharedApplication = new Excel.Application
                {
                    DisplayAlerts = true,
                    ScreenUpdating = true,
                    EnableEvents = true
                };
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                kernelWorkbook.Application = sharedApplication;
                kernelWorkbook.Worksheets.Add(CreateUserDataWorksheet());
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                InvalidOperationException expected = new InvalidOperationException("write failed");

                var service = new AccountingSetKernelSyncService(
                    new ExcelInteropService(),
                    new AccountingTemplateResolver
                    {
                        OnResolveTemplatePath = _ => templatePath
                    },
                    new AccountingWorkbookService
                    {
                        OnOpenInCurrentApplication = _ =>
                        {
                            accountingWorkbook.Application = sharedApplication;
                            return accountingWorkbook;
                        },
                        OnWriteCell = (_, __, ___, ____) => throw expected
                    },
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(new List<string>()));

                InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() => service.Execute(kernelWorkbook));

                Assert.Same(expected, actual);
                Assert.True(sharedApplication.DisplayAlerts);
                Assert.True(sharedApplication.ScreenUpdating);
                Assert.True(sharedApplication.EnableEvents);
                Assert.Equal(0, accountingWorkbook.SaveCallCount);
                Assert.Equal(1, accountingWorkbook.CloseCallCount);
                Assert.Equal(0, sharedApplication.QuitCallCount);
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenCurrentApplicationFallbackWriteFailsAndCloseFails_RethrowsOriginalWriteException()
        {
            List<string> logs = new List<string>();
            string tempDirectory = CreateTempDirectory();
            try
            {
                string templatePath = Path.Combine(tempDirectory, "accounting-template.xlsx");
                File.WriteAllText(templatePath, "template");
                Excel.Application sharedApplication = new Excel.Application
                {
                    DisplayAlerts = true,
                    ScreenUpdating = true,
                    EnableEvents = true
                };
                Excel.Workbook kernelWorkbook = CreateWorkbook(Path.Combine(tempDirectory, "kernel.xlsx"));
                kernelWorkbook.Application = sharedApplication;
                kernelWorkbook.Worksheets.Add(CreateUserDataWorksheet());
                Excel.Workbook accountingWorkbook = CreateWorkbook(templatePath);
                InvalidOperationException expected = new InvalidOperationException("write failed");
                accountingWorkbook.CloseBehavior = () => throw new InvalidOperationException("close failed " + templatePath);

                var service = new AccountingSetKernelSyncService(
                    new ExcelInteropService(),
                    new AccountingTemplateResolver
                    {
                        OnResolveTemplatePath = _ => templatePath
                    },
                    new AccountingWorkbookService
                    {
                        OnOpenInCurrentApplication = _ =>
                        {
                            accountingWorkbook.Application = sharedApplication;
                            return accountingWorkbook;
                        },
                        OnWriteCell = (_, __, ___, ____) => throw expected
                    },
                    new PathCompatibilityService(),
                    OrchestrationTestSupport.CreateLogger(logs));

                InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() => service.Execute(kernelWorkbook));

                Assert.Same(expected, actual);
                Assert.Equal(0, accountingWorkbook.SaveCallCount);
                Assert.Equal(1, accountingWorkbook.CloseCallCount);
                Assert.True(sharedApplication.DisplayAlerts);
                Assert.True(sharedApplication.ScreenUpdating);
                Assert.True(sharedApplication.EnableEvents);
                Assert.Contains(logs, message =>
                    ContainsOrdinal(message, "WARN: Accounting set kernel sync owned workbook close attempted.")
                    && ContainsOrdinal(message, "closeSucceeded=False")
                    && ContainsOrdinal(message, "exceptionType=System.InvalidOperationException"));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, tempDirectory));
                Assert.DoesNotContain(logs, message => ContainsOrdinal(message, "close failed"));
            }
            finally
            {
                TryDeleteDirectory(tempDirectory);
            }
        }

        [Fact]
        public void Execute_WhenUserDataWorksheetIsMissing_ThrowsInvalidOperationException()
        {
            var service = new AccountingSetKernelSyncService(
                new ExcelInteropService(),
                new AccountingTemplateResolver
                {
                    OnResolveTemplatePath = _ => "ignored"
                },
                new AccountingWorkbookService(),
                new PathCompatibilityService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()));

            Assert.Throws<InvalidOperationException>(() => service.Execute(CreateWorkbook(@"C:\temp\kernel.xlsx")));
        }

        private static bool ContainsOrdinal(string text, string value)
        {
            return (text ?? string.Empty).IndexOf(value ?? string.Empty, StringComparison.Ordinal) >= 0;
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
            Excel.Worksheet worksheet = new Excel.Worksheet
            {
                CodeName = AccountingSetSpec.UserDataSheetCodeName,
                Name = AccountingSetSpec.UserDataSheetName
            };
            worksheet.Cells.SetValue(AccountingSetSpec.UserDataFirstDataRow + 0, "B", 0);
            worksheet.Cells.SetValue(AccountingSetSpec.UserDataFirstDataRow + 1, "B", "東京都");
            worksheet.Cells.SetValue(AccountingSetSpec.UserDataFirstDataRow + 2, "B", "千代田区1-1");
            worksheet.Cells.SetValue(AccountingSetSpec.UserDataFirstDataRow + 3, "B", "03-1111-2222");
            worksheet.Cells.SetValue(AccountingSetSpec.UserDataFirstDataRow + AccountingSetSpec.UserDataAccountingNameRow1Offset, "B", "銀行支店ABC");
            worksheet.Cells.SetValue(AccountingSetSpec.UserDataFirstDataRow + AccountingSetSpec.UserDataAccountingNameRow2Offset, "B", "名義XYZ");
            return worksheet;
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
