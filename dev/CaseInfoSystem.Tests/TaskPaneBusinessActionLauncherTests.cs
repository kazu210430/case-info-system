using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class TaskPaneBusinessActionLauncherTests
    {
        [Fact]
        public void TryExecute_WhenDocumentNamePromptIsCancelled_ReturnsFalseAndSkipsCommandExecution()
        {
            bool commandExecuted = false;
            var promptService = new DocumentNamePromptService
            {
                OnTryPrepare = (workbook, key) => false
            };
            var service = new TaskPaneBusinessActionLauncher(
                CreateDocumentCommandService((workbook, templateSpec, caseContext) => commandExecuted = true),
                promptService);

            bool result = service.TryExecute(new Excel.Workbook(), "doc", "01");

            Assert.False(result);
            Assert.False(commandExecuted);
        }

        [Fact]
        public void TryExecute_WhenDocumentNamePromptIsAccepted_ExecutesDocumentCommand()
        {
            bool promptCalled = false;
            bool commandExecuted = false;
            var promptService = new DocumentNamePromptService
            {
                OnTryPrepare = (workbook, key) =>
                {
                    promptCalled = true;
                    return true;
                }
            };
            var service = new TaskPaneBusinessActionLauncher(
                CreateDocumentCommandService((workbook, templateSpec, caseContext) => commandExecuted = true),
                promptService);

            bool result = service.TryExecute(new Excel.Workbook(), "doc", "01");

            Assert.True(result);
            Assert.True(promptCalled);
            Assert.True(commandExecuted);
        }

        private static DocumentCommandService CreateDocumentCommandService(
            System.Action<Excel.Workbook, DocumentTemplateSpec, CaseContext> onDocumentCreate)
        {
            var documentExecutionEligibilityService = new DocumentExecutionEligibilityService
            {
                OnEvaluate = (workbook, actionKind, key) => new DocumentExecutionEligibility(
                    canExecuteInVsto: true,
                    reason: string.Empty,
                    templateSpec: new DocumentTemplateSpec
                    {
                        TemplateFileName = "01_委任状.docx"
                    },
                    caseContext: new CaseContext())
            };
            var documentCreateService = new DocumentCreateService
            {
                OnExecute = onDocumentCreate
            };
            return new DocumentCommandService(
                new CaseInfoSystem.ExcelAddIn.ThisAddIn(),
                new InlineScreenUpdatingExecutionBridge(),
                new NoOpTaskPaneRefreshSuppressionBridge(),
                new CollectingActiveTaskPaneRefreshBridge(),
                new DocumentExecutionModeService(OrchestrationTestSupport.CreateLogger(new List<string>()), new ExcelInteropService()),
                documentExecutionEligibilityService,
                documentCreateService,
                new AccountingSetCommandService(),
                new CaseListRegistrationService(),
                new CaseContextFactory(),
                new ExcelInteropService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()));
        }

        private sealed class InlineScreenUpdatingExecutionBridge : IScreenUpdatingExecutionBridge
        {
            public void Execute(System.Action action)
            {
                action?.Invoke();
            }
        }

        private sealed class NoOpTaskPaneRefreshSuppressionBridge : ITaskPaneRefreshSuppressionBridge
        {
            public System.IDisposable Enter(string reason)
            {
                return new NoOpDisposable();
            }
        }

        private sealed class CollectingActiveTaskPaneRefreshBridge : IActiveTaskPaneRefreshBridge
        {
            public void RequestRefresh(string reason)
            {
            }
        }

        private sealed class NoOpDisposable : System.IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
