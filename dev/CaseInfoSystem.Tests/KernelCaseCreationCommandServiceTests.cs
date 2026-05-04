using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class KernelCaseCreationCommandServiceTests
    {
        [Fact]
        public void ExecuteCreateCaseSingle_WhenBoundWorkbookIsNull_FailsClosedWithoutOpenKernelFallback()
        {
            int openKernelCalls = 0;
            KernelCaseCreationCommandService service = CreateService(
                new KernelWorkbookService(
                    OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    new KernelWorkbookService.KernelWorkbookServiceTestHooks
                    {
                    }),
                new KernelCasePathService
                {
                    OnResolveSystemRoot = workbook => workbook == null ? string.Empty : workbook.Path
                });

            var result = service.ExecuteCreateCaseSingle(null, @"C:\root", "ClientA");

            Assert.False(result.Success);
            Assert.NotEmpty(result.UserMessage);
            Assert.Equal(0, openKernelCalls);
        }

        [Fact]
        public void ExecuteCreateCaseSingle_WhenBoundWorkbookRootMismatches_FailsClosedWithoutOpenKernelFallback()
        {
            int openKernelCalls = 0;
            Excel.Workbook kernelWorkbook = new Excel.Workbook
            {
                Name = WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsm"),
                FullName = @"C:\actual-root\kernel.xlsm",
                Path = @"C:\actual-root"
            };
            KernelCaseCreationCommandService service = CreateService(
                new KernelWorkbookService(
                    OrchestrationTestSupport.CreateKernelCaseInteractionState(new List<string>()),
                    OrchestrationTestSupport.CreateLogger(new List<string>()),
                    new KernelWorkbookService.KernelWorkbookServiceTestHooks
                    {
                    }),
                new KernelCasePathService
                {
                    OnResolveSystemRoot = workbook => workbook == null ? string.Empty : workbook.Path
                });

            var result = service.ExecuteCreateCaseSingle(kernelWorkbook, @"C:\expected-root", "ClientA");

            Assert.False(result.Success);
            Assert.NotEmpty(result.UserMessage);
            Assert.Equal(0, openKernelCalls);
        }

        private static KernelCaseCreationCommandService CreateService(KernelWorkbookService kernelWorkbookService, KernelCasePathService kernelCasePathService)
        {
            Logger logger = OrchestrationTestSupport.CreateLogger(new List<string>());
            return new KernelCaseCreationCommandService(
                kernelWorkbookService,
                CreateUnused<KernelCaseCreationService>(),
                kernelCasePathService,
                CreateUnused<KernelCasePresentationService>(),
                CreateUnused<CreatedCasePresentationWaitService>(),
                CreateUnused<CaseWorkbookLifecycleService>(),
                new ExcelInteropService(),
                logger);
        }

        private static T CreateUnused<T>() where T : class
        {
            return (T)FormatterServices.GetUninitializedObject(typeof(T));
        }
    }
}
