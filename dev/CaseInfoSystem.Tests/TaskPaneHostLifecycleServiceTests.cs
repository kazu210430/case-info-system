using System;
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
    public class TaskPaneHostLifecycleServiceTests
    {
        [Fact]
        public void ResolveRefreshHost_RemovesOtherKernelHostsForSameWorkbookBeforeSelectingActiveHost()
        {
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
            TaskPaneHostRegistry registry = CreateRegistry(hostsByWindowKey);
            var service = new TaskPaneHostLifecycleService(
                hostsByWindowKey,
                registry,
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                host => host == null ? string.Empty : host.WindowKey);

            TaskPaneHost activeHost = OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "101");
            activeHost.WorkbookFullName = @"C:\kernel\kernel.xlsx";
            hostsByWindowKey["101"] = activeHost;

            TaskPaneHost staleHost = OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "202");
            staleHost.WorkbookFullName = @"C:\kernel\kernel.xlsx";
            hostsByWindowKey["202"] = staleHost;

            TaskPaneHost otherWorkbookHost = OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "303");
            otherWorkbookHost.WorkbookFullName = @"C:\kernel\other.xlsx";
            hostsByWindowKey["303"] = otherWorkbookHost;

            var context = new WorkbookContext(
                new Excel.Workbook { FullName = @"C:\kernel\kernel.xlsx", Name = "kernel.xlsx" },
                new Excel.Window { Hwnd = 101 },
                WorkbookRole.Kernel,
                @"C:\kernel",
                @"C:\kernel\kernel.xlsx",
                "shHOME");

            TaskPaneHost resolvedHost = service.ResolveRefreshHost(context, "101", refreshPaneCallId: 1);

            Assert.Same(activeHost, resolvedHost);
            Assert.True(hostsByWindowKey.ContainsKey("101"));
            Assert.False(hostsByWindowKey.ContainsKey("202"));
            Assert.True(hostsByWindowKey.ContainsKey("303"));
        }

        [Fact]
        public void RemoveWorkbookPanes_RemovesOnlyHostsForTargetWorkbook()
        {
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
            TaskPaneHostRegistry registry = CreateRegistry(hostsByWindowKey);
            var service = new TaskPaneHostLifecycleService(
                hostsByWindowKey,
                registry,
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                host => host == null ? string.Empty : host.WindowKey);

            TaskPaneHost caseHost = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "101");
            caseHost.WorkbookFullName = @"C:\cases\target.xlsx";
            hostsByWindowKey["101"] = caseHost;

            TaskPaneHost accountingHost = OrchestrationTestSupport.CreateTaskPaneHost(new AccountingNavigationControl(), "202");
            accountingHost.WorkbookFullName = @"C:\cases\target.xlsx";
            hostsByWindowKey["202"] = accountingHost;

            TaskPaneHost otherHost = OrchestrationTestSupport.CreateTaskPaneHost(new DocumentButtonsControl(), "303");
            otherHost.WorkbookFullName = @"C:\cases\other.xlsx";
            hostsByWindowKey["303"] = otherHost;

            service.RemoveWorkbookPanes(new Excel.Workbook { FullName = @"C:\cases\target.xlsx", Name = "target.xlsx" });

            Assert.False(hostsByWindowKey.ContainsKey("101"));
            Assert.False(hostsByWindowKey.ContainsKey("202"));
            Assert.True(hostsByWindowKey.ContainsKey("303"));
        }

        private static TaskPaneHostRegistry CreateRegistry(Dictionary<string, TaskPaneHost> hostsByWindowKey)
        {
            return new TaskPaneHostRegistry(
                hostsByWindowKey,
                new CaseInfoSystem.ExcelAddIn.ThisAddIn(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                host => host == null ? string.Empty : host.WindowKey,
                (windowKey, e) => { },
                (windowKey, e) => { },
                (windowKey, control, e) => { });
        }
    }
}
