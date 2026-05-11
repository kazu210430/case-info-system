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
    public class TaskPaneActionDispatcherTests
    {
        [Fact]
        public void HandleCaseControlActionInvoked_WhenDocumentPromptIsCancelled_DoesNotRefresh()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            var excelInteropService = new ExcelInteropService();
            Excel.Workbook workbook = CreateWorkbook(@"C:\cases\case.xlsx");
            excelInteropService.OnFindOpenWorkbook = _ => workbook;

            var control = new DocumentButtonsControl();
            TaskPaneHost host = OrchestrationTestSupport.CreateTaskPaneHost(control, "101");
            host.WorkbookFullName = workbook.FullName;
            int invalidateCalls = 0;
            int renderAfterActionCalls = 0;
            int showCalls = 0;
            int requestCalls = 0;
            addIn.RequestTaskPaneDisplayForTargetWindowHandler = (request, targetWorkbook, targetWindow) => requestCalls++;

            var dispatcher = CreateDispatcher(
                addIn,
                excelInteropService,
                CreateBusinessActionLauncher(
                    promptAccepted: false,
                    onDocumentCreate: null,
                    onAccountingExecute: null,
                    caseListResultFactory: null,
                    caseListContextFactory: null),
                new CaseTaskPaneViewStateBuilder(),
                new UserErrorService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                windowKey => host,
                _ => invalidateCalls++,
                (targetControl, targetWorkbook) => renderAfterActionCalls++,
                (targetHost, reason) =>
                {
                    showCalls++;
                    return true;
                });

            dispatcher.HandleCaseControlActionInvoked("101", control, new TaskPaneActionEventArgs("doc", "01"));

            Assert.Equal(0, requestCalls);
            Assert.Equal(0, invalidateCalls);
            Assert.Equal(0, renderAfterActionCalls);
            Assert.Equal(0, showCalls);
        }

        [Fact]
        public void HandleCaseControlActionInvoked_WhenDocumentActionRuns_SkipsPostActionRefresh()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            var excelInteropService = new ExcelInteropService();
            Excel.Workbook workbook = CreateWorkbook(@"C:\cases\case.xlsx");
            excelInteropService.OnFindOpenWorkbook = _ => workbook;

            var control = new DocumentButtonsControl();
            TaskPaneHost host = OrchestrationTestSupport.CreateTaskPaneHost(control, "202");
            host.WorkbookFullName = workbook.FullName;
            bool documentCreated = false;
            int invalidateCalls = 0;
            int renderAfterActionCalls = 0;
            int showCalls = 0;
            int requestCalls = 0;
            addIn.RequestTaskPaneDisplayForTargetWindowHandler = (request, targetWorkbook, targetWindow) => requestCalls++;

            var dispatcher = CreateDispatcher(
                addIn,
                excelInteropService,
                CreateBusinessActionLauncher(
                    promptAccepted: true,
                    onDocumentCreate: (targetWorkbook, templateSpec, caseContext) => documentCreated = true,
                    onAccountingExecute: null,
                    caseListResultFactory: null,
                    caseListContextFactory: null),
                new CaseTaskPaneViewStateBuilder(),
                new UserErrorService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                windowKey => host,
                _ => invalidateCalls++,
                (targetControl, targetWorkbook) => renderAfterActionCalls++,
                (targetHost, reason) =>
                {
                    showCalls++;
                    return true;
                });

            dispatcher.HandleCaseControlActionInvoked("202", control, new TaskPaneActionEventArgs("doc", "01"));

            Assert.True(documentCreated);
            Assert.Equal(0, requestCalls);
            Assert.Equal(0, invalidateCalls);
            Assert.Equal(0, renderAfterActionCalls);
            Assert.Equal(0, showCalls);
        }

        [Fact]
        public void HandleCaseControlActionInvoked_WhenAccountingActionRuns_SkipsPostActionRefresh()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            var excelInteropService = new ExcelInteropService();
            Excel.Workbook workbook = CreateWorkbook(@"C:\cases\case.xlsx");
            excelInteropService.OnFindOpenWorkbook = _ => workbook;

            var control = new DocumentButtonsControl();
            TaskPaneHost host = OrchestrationTestSupport.CreateTaskPaneHost(control, "252");
            host.WorkbookFullName = workbook.FullName;
            bool accountingExecuted = false;
            int invalidateCalls = 0;
            int renderAfterActionCalls = 0;
            int showCalls = 0;
            int requestCalls = 0;
            addIn.RequestTaskPaneDisplayForTargetWindowHandler = (request, targetWorkbook, targetWindow) => requestCalls++;

            var dispatcher = CreateDispatcher(
                addIn,
                excelInteropService,
                CreateBusinessActionLauncher(
                    promptAccepted: true,
                    onDocumentCreate: null,
                    onAccountingExecute: _ => accountingExecuted = true,
                    caseListResultFactory: null,
                    caseListContextFactory: null),
                new CaseTaskPaneViewStateBuilder(),
                new UserErrorService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                windowKey => host,
                _ => invalidateCalls++,
                (targetControl, targetWorkbook) => renderAfterActionCalls++,
                (targetHost, reason) =>
                {
                    showCalls++;
                    return true;
                });

            dispatcher.HandleCaseControlActionInvoked("252", control, new TaskPaneActionEventArgs("accounting", string.Empty));

            Assert.True(accountingExecuted);
            Assert.Equal(0, requestCalls);
            Assert.Equal(0, invalidateCalls);
            Assert.Equal(0, renderAfterActionCalls);
            Assert.Equal(0, showCalls);
        }

        [Fact]
        public void HandleCaseControlActionInvoked_WhenAccountingActionThrows_LogsErrorAndSkipsPostActionRefresh()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            var excelInteropService = new ExcelInteropService();
            Excel.Workbook workbook = CreateWorkbook(@"C:\cases\case.xlsx");
            excelInteropService.OnFindOpenWorkbook = _ => workbook;

            var control = new DocumentButtonsControl();
            TaskPaneHost host = OrchestrationTestSupport.CreateTaskPaneHost(control, "262");
            host.WorkbookFullName = workbook.FullName;
            var expectedException = new System.InvalidOperationException("Accounting action failed.");
            var loggerMessages = new List<string>();
            string userErrorContext = null;
            System.Exception userErrorException = null;
            int invalidateCalls = 0;
            int renderAfterActionCalls = 0;
            int showCalls = 0;
            int requestCalls = 0;
            addIn.RequestTaskPaneDisplayForTargetWindowHandler = (request, targetWorkbook, targetWindow) => requestCalls++;
            var userErrorService = new UserErrorService
            {
                OnShowUserError = (context, ex) =>
                {
                    userErrorContext = context;
                    userErrorException = ex;
                }
            };

            var dispatcher = CreateDispatcher(
                addIn,
                excelInteropService,
                CreateBusinessActionLauncher(
                    promptAccepted: true,
                    onDocumentCreate: null,
                    onAccountingExecute: _ => throw expectedException,
                    caseListResultFactory: null,
                    caseListContextFactory: null),
                new CaseTaskPaneViewStateBuilder(),
                userErrorService,
                OrchestrationTestSupport.CreateLogger(loggerMessages),
                windowKey => host,
                _ => invalidateCalls++,
                (targetControl, targetWorkbook) => renderAfterActionCalls++,
                (targetHost, reason) =>
                {
                    showCalls++;
                    return true;
                });

            dispatcher.HandleCaseControlActionInvoked("262", control, new TaskPaneActionEventArgs("accounting", string.Empty));

            Assert.Contains(
                loggerMessages,
                message => message.Contains("ERROR: CaseControl_ActionInvoked failed.")
                    && message.Contains("InvalidOperationException")
                    && message.Contains("Accounting action failed."));
            Assert.Equal("CaseControl_ActionInvoked", userErrorContext);
            Assert.Same(expectedException, userErrorException);
            Assert.Equal(0, requestCalls);
            Assert.Equal(0, invalidateCalls);
            Assert.Equal(0, renderAfterActionCalls);
            Assert.Equal(0, showCalls);
        }

        [Fact]
        public void HandleCaseControlActionInvoked_WhenCaseListActionRuns_SkipsRefreshAndKeepsSignature()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            var excelInteropService = new ExcelInteropService();
            Excel.Workbook workbook = CreateWorkbook(@"C:\cases\case.xlsx");
            excelInteropService.OnFindOpenWorkbook = _ => workbook;

            var control = new DocumentButtonsControl();
            TaskPaneHost host = OrchestrationTestSupport.CreateTaskPaneHost(control, "303");
            host.WorkbookFullName = workbook.FullName;
            host.LastRenderSignature = "existing-signature";
            int invalidateCalls = 0;
            int renderAfterActionCalls = 0;
            int showCalls = 0;
            int requestCalls = 0;
            addIn.RequestTaskPaneDisplayForTargetWindowHandler = (request, targetWorkbook, targetWindow) => requestCalls++;

            var dispatcher = CreateDispatcher(
                addIn,
                excelInteropService,
                CreateBusinessActionLauncher(
                    promptAccepted: true,
                    onDocumentCreate: null,
                    onAccountingExecute: null,
                    caseListResultFactory: () => new CaseListRegistrationResult
                    {
                        Success = true,
                        RegisteredRow = 5,
                        Message = "registered"
                    },
                    caseListContextFactory: targetWorkbook => null),
                new CaseTaskPaneViewStateBuilder(),
                new UserErrorService(),
                OrchestrationTestSupport.CreateLogger(new List<string>()),
                windowKey => host,
                _ => invalidateCalls++,
                (targetControl, targetWorkbook) => renderAfterActionCalls++,
                (targetHost, reason) =>
                {
                    showCalls++;
                    return true;
                });

            dispatcher.HandleCaseControlActionInvoked("303", control, new TaskPaneActionEventArgs("caselist", string.Empty));

            Assert.Equal(0, requestCalls);
            Assert.Equal(0, invalidateCalls);
            Assert.Equal(0, renderAfterActionCalls);
            Assert.Equal(0, showCalls);
            Assert.Equal("existing-signature", host.LastRenderSignature);
        }

        private static Excel.Workbook CreateWorkbook(string fullName)
        {
            return new Excel.Workbook
            {
                FullName = fullName,
                Name = "case.xlsx",
                ActiveSheet = new Excel.Worksheet { CodeName = "shHOME" },
                CustomDocumentProperties = new Dictionary<string, string>
                {
                    ["SYSTEM_ROOT"] = @"C:\cases"
                }
            };
        }

        private static TaskPaneBusinessActionLauncher CreateBusinessActionLauncher(
            bool promptAccepted,
            System.Action<Excel.Workbook, DocumentTemplateSpec, CaseContext> onDocumentCreate,
            System.Action<Excel.Workbook> onAccountingExecute,
            System.Func<CaseListRegistrationResult> caseListResultFactory,
            System.Func<Excel.Workbook, CaseContext> caseListContextFactory)
        {
            var promptService = new DocumentNamePromptService
            {
                OnTryPrepare = (workbook, key) => promptAccepted
            };
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
            var accountingSetCommandService = new AccountingSetCommandService
            {
                OnExecute = onAccountingExecute
            };
            var caseListRegistrationService = new CaseListRegistrationService
            {
                OnExecute = workbook => caseListResultFactory == null
                    ? new CaseListRegistrationResult { Success = true }
                    : caseListResultFactory()
            };
            var caseContextFactory = new CaseContextFactory
            {
                OnCreateForCaseListRegistration = caseListContextFactory
            };
            return new TaskPaneBusinessActionLauncher(
                new DocumentCommandService(
                    new CaseInfoSystem.ExcelAddIn.ThisAddIn(),
                    new InlineScreenUpdatingExecutionBridge(),
                    new NoOpTaskPaneRefreshSuppressionBridge(),
                    new CollectingActiveTaskPaneRefreshBridge(),
                    new DocumentExecutionModeService(OrchestrationTestSupport.CreateLogger(new List<string>()), new ExcelInteropService()),
                    documentExecutionEligibilityService,
                    documentCreateService,
                    accountingSetCommandService,
                    caseListRegistrationService,
                    caseContextFactory,
                    new ExcelInteropService(),
                    OrchestrationTestSupport.CreateLogger(new List<string>())),
                promptService);
        }

        private static TaskPaneActionDispatcher CreateDispatcher(
            CaseInfoSystem.ExcelAddIn.ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            System.Func<string, TaskPaneHost> resolveHost,
            System.Action<TaskPaneHost> invalidateHostRenderStateForForcedRefresh,
            System.Action<DocumentButtonsControl, Excel.Workbook> renderCaseHostAfterAction,
            System.Func<TaskPaneHost, string, bool> tryShowHost)
        {
            var taskPaneCaseFallbackActionExecutor = new TaskPaneCaseFallbackActionExecutor(taskPaneBusinessActionLauncher);
            var taskPaneCaseActionTargetResolver = new TaskPaneCaseActionTargetResolver(
                excelInteropService,
                logger,
                resolveHost);
            TaskPaneActionDispatcher dispatcher = null;
            System.Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh =
                (host, workbook, control, actionKind) => dispatcher.HandlePostActionRefresh(host, workbook, control, actionKind);
            var taskPaneCaseAccountingActionHandler = new TaskPaneCaseAccountingActionHandler(
                taskPaneCaseActionTargetResolver,
                taskPaneCaseFallbackActionExecutor,
                caseTaskPaneViewStateBuilder,
                userErrorService,
                logger,
                handlePostActionRefresh);
            var taskPaneCaseDocumentActionHandler = new TaskPaneCaseDocumentActionHandler(
                taskPaneCaseActionTargetResolver,
                taskPaneCaseFallbackActionExecutor,
                caseTaskPaneViewStateBuilder,
                userErrorService,
                logger,
                handlePostActionRefresh);
            dispatcher = new TaskPaneActionDispatcher(
                addIn,
                excelInteropService,
                caseTaskPaneViewStateBuilder,
                userErrorService,
                logger,
                taskPaneCaseFallbackActionExecutor,
                taskPaneCaseActionTargetResolver,
                taskPaneCaseAccountingActionHandler,
                taskPaneCaseDocumentActionHandler,
                invalidateHostRenderStateForForcedRefresh,
                renderCaseHostAfterAction,
                tryShowHost);
            return dispatcher;
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
