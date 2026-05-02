using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.UI;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneHostFactoryTests
    {
        [Theory]
        [InlineData((int)WorkbookRole.Kernel, typeof(KernelNavigationControl), "Kernel")]
        [InlineData((int)WorkbookRole.Accounting, typeof(AccountingNavigationControl), "Accounting")]
        [InlineData((int)WorkbookRole.Case, typeof(DocumentButtonsControl), "Case")]
        [InlineData((int)WorkbookRole.Unknown, typeof(DocumentButtonsControl), "Case")]
        public void CreateHost_CreatesExpectedHostForRole(int roleValue, Type controlType, string expectedPaneRoleName)
        {
            var factory = CreateFactory();
            WorkbookRole role = (WorkbookRole)roleValue;

            TaskPaneHost host = factory.CreateHost("101", new Excel.Window { Hwnd = 101 }, role, out string paneRoleName);

            Assert.NotNull(host);
            Assert.IsType(controlType, host.Control);
            Assert.Equal("101", host.WindowKey);
            Assert.Equal(expectedPaneRoleName, paneRoleName);
        }

        [Fact]
        public void GetOrReplaceHost_ReturnsExistingHost_WhenRoleMatches()
        {
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
            var registry = CreateRegistry(hostsByWindowKey);
            TaskPaneHost existingHost = OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "101");
            hostsByWindowKey["101"] = existingHost;

            TaskPaneHost resolvedHost = registry.GetOrReplaceHost("101", new Excel.Window { Hwnd = 101 }, WorkbookRole.Kernel);

            Assert.Same(existingHost, resolvedHost);
        }

        [Fact]
        public void GetOrReplaceHost_ReplacesExistingHost_WhenRoleDoesNotMatch()
        {
            var hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
            var registry = CreateRegistry(hostsByWindowKey);
            TaskPaneHost existingHost = OrchestrationTestSupport.CreateTaskPaneHost(new KernelNavigationControl(), "101");
            hostsByWindowKey["101"] = existingHost;

            TaskPaneHost resolvedHost = registry.GetOrReplaceHost("101", new Excel.Window { Hwnd = 101 }, WorkbookRole.Case);

            Assert.NotSame(existingHost, resolvedHost);
            Assert.IsType<DocumentButtonsControl>(resolvedHost.Control);
            Assert.Same(resolvedHost, hostsByWindowKey["101"]);
        }

        private static TaskPaneHostFactory CreateFactory()
        {
            return new TaskPaneHostFactory(
                new CaseInfoSystem.ExcelAddIn.ThisAddIn(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                host => host == null ? string.Empty : host.WindowKey,
                (windowKey, e) => { },
                (windowKey, e) => { },
                (windowKey, control, e) => { });
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
