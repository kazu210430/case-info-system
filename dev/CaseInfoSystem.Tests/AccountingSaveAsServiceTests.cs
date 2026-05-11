using System;
using System.Collections.Generic;
using System.Reflection;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class AccountingSaveAsServiceTests
    {
        [Fact]
        public void TryDisableAutoSave_WhenAutoSaveIsAvailable_TurnsOffWorkbookAutoSave()
        {
            List<string> logs = new List<string>();
            AccountingSaveAsService service = CreateService(logs);
            Excel.Workbook workbook = CreateWorkbook();
            workbook.AutoSaveOn = true;

            InvokeTryDisableAutoSave(service, workbook, @"C:\work\accounting-set.xlsm");

            Assert.False(workbook.AutoSaveOn);
            Assert.Contains(logs, message => message.Contains("Accounting workbook AutoSave disabled after SaveAs."));
        }

        [Fact]
        public void TryDisableAutoSave_WhenAutoSavePropertyFails_LogsWarningAndContinues()
        {
            List<string> logs = new List<string>();
            AccountingSaveAsService service = CreateService(logs);
            Excel.Workbook workbook = CreateWorkbook();
            workbook.AutoSaveOn = true;
            workbook.AutoSaveOnSetBehavior = _ => throw new InvalidOperationException("autosave unavailable");

            Exception exception = Record.Exception(() => InvokeTryDisableAutoSave(service, workbook, @"C:\work\accounting-set.xlsm"));

            Assert.Null(exception);
            Assert.True(workbook.AutoSaveOn);
            Assert.Contains(logs, message =>
                message.Contains("WARN: Accounting workbook AutoSave disable after SaveAs skipped.")
                && message.Contains("autosave unavailable"));
        }

        private static AccountingSaveAsService CreateService(List<string> logs)
        {
            Logger logger = OrchestrationTestSupport.CreateLogger(logs);
            ExcelInteropService excelInteropService = new ExcelInteropService();
            PathCompatibilityService pathCompatibilityService = new PathCompatibilityService(logger);
            return new AccountingSaveAsService(
                excelInteropService,
                new AccountingWorkbookService(),
                new DocumentOutputService(excelInteropService, pathCompatibilityService, logger),
                pathCompatibilityService,
                new UserErrorService(),
                logger);
        }

        private static Excel.Workbook CreateWorkbook()
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = new Excel.Workbook
            {
                Application = application,
                FullName = @"C:\work\accounting-set.xlsx",
                Name = "accounting-set.xlsx",
                Path = @"C:\work"
            };
            application.Workbooks.Add(workbook);
            return workbook;
        }

        private static void InvokeTryDisableAutoSave(AccountingSaveAsService service, Excel.Workbook workbook, string savePath)
        {
            MethodInfo method = typeof(AccountingSaveAsService).GetMethod("TryDisableAutoSave", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(method);
            method.Invoke(service, new object[] { workbook, savePath });
        }
    }
}
