using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneRenderStateEvaluatorTests
    {
        [Fact]
        public void EvaluateDisplayEntryState_WhenSameWorkbookAndSignatureMatches_ReturnsCurrent()
        {
            var excelInteropService = new ExcelInteropService();
            var workbook = CreateCaseWorkbook("shHOME", caseListRegistered: "0", snapshotCacheCount: "2");
            var window = new Excel.Window { Hwnd = 101 };
            var host = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "101");
            host.WorkbookFullName = workbook.FullName;
            host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(
                excelInteropService,
                CreateCaseContext(workbook, window));
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["101"] = host
            };

            TaskPaneDisplayEntryState result = TaskPaneRenderStateEvaluator.EvaluateDisplayEntryState(
                excelInteropService,
                hostsByWindowKey,
                workbook,
                window);

            Assert.True(result.HasTargetWindow);
            Assert.True(result.HasResolvableWindowKey);
            Assert.True(result.HasManagedPane);
            Assert.True(result.HasExistingHost);
            Assert.True(result.IsSameWorkbook);
            Assert.True(result.IsRenderSignatureCurrent);
        }

        [Fact]
        public void EvaluateDisplayEntryState_WhenActiveSheetChanges_ReturnsSignatureOutdated()
        {
            var excelInteropService = new ExcelInteropService();
            var workbook = CreateCaseWorkbook("shHOME", caseListRegistered: "0", snapshotCacheCount: "2");
            var window = new Excel.Window { Hwnd = 101 };
            var host = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "101");
            host.WorkbookFullName = workbook.FullName;
            host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(
                excelInteropService,
                CreateCaseContext(workbook, window));
            workbook.ActiveSheet = new Excel.Worksheet { CodeName = "shDETAIL" };
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["101"] = host
            };

            TaskPaneDisplayEntryState result = TaskPaneRenderStateEvaluator.EvaluateDisplayEntryState(
                excelInteropService,
                hostsByWindowKey,
                workbook,
                window);

            Assert.True(result.HasTargetWindow);
            Assert.True(result.HasResolvableWindowKey);
            Assert.True(result.HasManagedPane);
            Assert.True(result.HasExistingHost);
            Assert.True(result.IsSameWorkbook);
            Assert.False(result.IsRenderSignatureCurrent);
        }

        [Fact]
        public void EvaluateDisplayEntryState_WhenCaseListRegisteredChanges_ReturnsSignatureOutdated()
        {
            var excelInteropService = new ExcelInteropService();
            var workbook = CreateCaseWorkbook("shHOME", caseListRegistered: "0", snapshotCacheCount: "2");
            var window = new Excel.Window { Hwnd = 101 };
            var host = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "101");
            host.WorkbookFullName = workbook.FullName;
            host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(
                excelInteropService,
                CreateCaseContext(workbook, window));
            ((IDictionary<string, string>)workbook.CustomDocumentProperties)["CASELIST_REGISTERED"] = "1";
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["101"] = host
            };

            TaskPaneDisplayEntryState result = TaskPaneRenderStateEvaluator.EvaluateDisplayEntryState(
                excelInteropService,
                hostsByWindowKey,
                workbook,
                window);

            Assert.True(result.HasTargetWindow);
            Assert.True(result.HasResolvableWindowKey);
            Assert.True(result.HasManagedPane);
            Assert.True(result.HasExistingHost);
            Assert.True(result.IsSameWorkbook);
            Assert.False(result.IsRenderSignatureCurrent);
        }

        [Fact]
        public void EvaluateDisplayEntryState_WhenSnapshotCacheCountChanges_ReturnsSignatureOutdated()
        {
            var excelInteropService = new ExcelInteropService();
            var workbook = CreateCaseWorkbook("shHOME", caseListRegistered: "0", snapshotCacheCount: "2");
            var window = new Excel.Window { Hwnd = 101 };
            var host = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "101");
            host.WorkbookFullName = workbook.FullName;
            host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(
                excelInteropService,
                CreateCaseContext(workbook, window));
            ((IDictionary<string, string>)workbook.CustomDocumentProperties)["TASKPANE_SNAPSHOT_CACHE_COUNT"] = "3";
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["101"] = host
            };

            TaskPaneDisplayEntryState result = TaskPaneRenderStateEvaluator.EvaluateDisplayEntryState(
                excelInteropService,
                hostsByWindowKey,
                workbook,
                window);

            Assert.True(result.HasTargetWindow);
            Assert.True(result.HasResolvableWindowKey);
            Assert.True(result.HasManagedPane);
            Assert.True(result.HasExistingHost);
            Assert.True(result.IsSameWorkbook);
            Assert.False(result.IsRenderSignatureCurrent);
        }

        [Fact]
        public void EvaluateRenderState_WhenLastRenderSignatureMatches_ReturnsRenderNotRequired()
        {
            var excelInteropService = new ExcelInteropService();
            var workbook = CreateCaseWorkbook("shHOME", caseListRegistered: "0", snapshotCacheCount: "2");
            var window = new Excel.Window { Hwnd = 101 };
            var context = CreateCaseContext(workbook, window);
            var host = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "101");
            host.LastRenderSignature = TaskPaneRenderStateEvaluator.BuildRenderSignature(excelInteropService, context);

            TaskPaneRenderStateEvaluation result = TaskPaneRenderStateEvaluator.EvaluateRenderState(
                excelInteropService,
                host,
                context);

            Assert.False(result.IsRenderRequired);
            Assert.Equal(host.LastRenderSignature, result.RenderSignature);
        }

        private static Excel.Workbook CreateCaseWorkbook(string activeSheetCodeName, string caseListRegistered, string snapshotCacheCount)
        {
            return new Excel.Workbook
            {
                FullName = @"C:\cases\case.xlsx",
                Name = "case.xlsx",
                ActiveSheet = new Excel.Worksheet { CodeName = activeSheetCodeName },
                CustomDocumentProperties = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = @"C:\cases",
                    ["CASELIST_REGISTERED"] = caseListRegistered,
                    ["TASKPANE_SNAPSHOT_CACHE_COUNT"] = snapshotCacheCount
                }
            };
        }

        private static WorkbookContext CreateCaseContext(Excel.Workbook workbook, Excel.Window window)
        {
            return new WorkbookContext(
                workbook,
                window,
                WorkbookRole.Case,
                @"C:\cases",
                workbook.FullName,
                workbook.ActiveSheet.CodeName);
        }
    }
}
