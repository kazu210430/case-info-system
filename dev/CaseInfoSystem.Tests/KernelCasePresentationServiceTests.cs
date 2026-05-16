using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public sealed class KernelCasePresentationServiceTests
    {
        [Fact]
        public void OpenCreatedCase_WhenPresentationStepsSucceed_MarksOutcomeCompleted()
        {
            var logs = new List<string>();
            Excel.Application application = CreateApplication();
            var bridge = new RecordingCasePaneHostBridge();
            KernelCasePresentationService service = CreateService(application, logs, bridge, recoverWithoutShowing: true);
            var waitSession = new CreatedCasePresentationWaitService.WaitSession();
            var result = CreateSuccessfulResult();

            KernelCaseCreationResult opened = service.OpenCreatedCase(result, waitSession);

            Assert.Same(result, opened);
            Assert.True(opened.Success);
            Assert.Equal(CasePresentationOutcome.Completed, opened.PresentationOutcome);
            Assert.NotNull(opened.CreatedWorkbook);
            Assert.True(opened.CreatedWorkbook.Windows[1].Visible);
            Assert.Same(opened.CreatedWorkbook, bridge.ReadyShowWorkbook);
            Assert.Equal(ControlFlowReasons.CreatedCasePostRelease, bridge.SuppressionReason);
            Assert.Equal(ControlFlowReasons.CreatedCasePostRelease, bridge.ReadyShowReason);
            Assert.True(waitSession.ClosedForSuccessfulPresentation);
            Assert.False(waitSession.ClosedAndRestoredOwner);
            Assert.Contains(logs, message => message.IndexOf("presentationOutcome=Completed", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        public void OpenCreatedCase_WhenReadyShowRequestThrows_MarksOutcomeDegradedWithoutClosingCreatedWorkbook()
        {
            var logs = new List<string>();
            Excel.Application application = CreateApplication();
            var bridge = new RecordingCasePaneHostBridge
            {
                ThrowOnReadyShow = true
            };
            KernelCasePresentationService service = CreateService(application, logs, bridge, recoverWithoutShowing: true);
            var waitSession = new CreatedCasePresentationWaitService.WaitSession();
            var result = CreateSuccessfulResult();

            KernelCaseCreationResult opened = service.OpenCreatedCase(result, waitSession);

            Assert.True(opened.Success);
            Assert.Equal(CasePresentationOutcome.Degraded, opened.PresentationOutcome);
            Assert.Contains("DeferredPresentationException", opened.PresentationOutcomeReason);
            Assert.NotNull(opened.CreatedWorkbook);
            Assert.Contains(opened.CreatedWorkbook, application.Workbooks);
            Assert.Equal(0, opened.CreatedWorkbook.CloseCallCount);
            Assert.True(waitSession.ClosedForSuccessfulPresentation);
            Assert.Contains(logs, message => message.IndexOf("ShowCreatedCase deferred presentation failed", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Contains(logs, message => message.IndexOf("presentationOutcome=Degraded", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        public void OpenCreatedCase_WhenWorkbookOpenThrows_MarksOutcomeFailedForCaller()
        {
            var logs = new List<string>();
            Excel.Application application = CreateApplication();
            application.Workbooks.OpenBehavior = (_, __, ___) => throw new InvalidOperationException("open failed");
            var bridge = new RecordingCasePaneHostBridge();
            KernelCasePresentationService service = CreateService(application, logs, bridge, recoverWithoutShowing: true);
            var waitSession = new CreatedCasePresentationWaitService.WaitSession();
            var result = CreateSuccessfulResult();

            Assert.Throws<InvalidOperationException>(() => service.OpenCreatedCase(result, waitSession));

            Assert.Equal(CasePresentationOutcome.Failed, result.PresentationOutcome);
            Assert.Contains("OpenCreatedCaseException", result.PresentationOutcomeReason);
            Assert.True(waitSession.ClosedAndRestoredOwner);
            Assert.True(waitSession.Disposed);
            Assert.Null(result.CreatedWorkbook);
            Assert.Contains(logs, message => message.IndexOf("presentationOutcome=Failed", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static KernelCasePresentationService CreateService(
            Excel.Application application,
            List<string> logs,
            RecordingCasePaneHostBridge bridge,
            bool recoverWithoutShowing)
        {
            Logger logger = OrchestrationTestSupport.CreateLogger(logs);
            var pathCompatibilityService = new PathCompatibilityService();
            var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);
            var recoveryService = new ExcelWindowRecoveryService(application, excelInteropService, logger)
            {
                OnTryRecoverWorkbookWindowWithoutShowing = (_, __, ___) => recoverWithoutShowing
            };
            var suppressionInterop = new FakeExcelInteropService
            {
                WorkbookFullName = @"C:\Cases\display.xlsx"
            };

            return new KernelCasePresentationService(
                application,
                new CaseWorkbookOpenStrategy(application, new WorkbookRoleResolver(), logger),
                excelInteropService,
                recoveryService,
                CreateUnused<KernelWorkbookResolverService>(),
                CreateUnused<CaseListFieldDefinitionRepository>(),
                new FolderWindowService(pathCompatibilityService, logger),
                new CreatedCasePresentationWaitService(logger),
                new TransientPaneSuppressionService(suppressionInterop, pathCompatibilityService, logger),
                bridge,
                new WorkbookWindowVisibilityService(excelInteropService, logger),
                logger);
        }

        private static Excel.Application CreateApplication()
        {
            return new Excel.Application
            {
                Visible = true,
                ScreenUpdating = true,
                EnableEvents = true,
                DisplayAlerts = true
            };
        }

        private static KernelCaseCreationResult CreateSuccessfulResult()
        {
            return new KernelCaseCreationResult
            {
                Success = true,
                Mode = KernelCaseCreationMode.NewCaseDefault,
                CaseFolderPath = @"C:\Cases",
                CaseWorkbookPath = @"C:\Cases\display.xlsx"
            };
        }

        private static T CreateUnused<T>() where T : class
        {
            return (T)FormatterServices.GetUninitializedObject(typeof(T));
        }

        private sealed class RecordingCasePaneHostBridge : ICasePaneHostBridge
        {
            internal bool ThrowOnReadyShow { get; set; }

            internal Excel.Workbook ReadyShowWorkbook { get; private set; }

            internal string ReadyShowReason { get; private set; }

            internal string SuppressionReason { get; private set; }

            public void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason)
            {
                SuppressionReason = reason ?? string.Empty;
            }

            public void ShowWorkbookTaskPaneWhenReady(Excel.Workbook workbook, string reason)
            {
                ReadyShowWorkbook = workbook;
                ReadyShowReason = reason ?? string.Empty;
                if (ThrowOnReadyShow)
                {
                    throw new InvalidOperationException("ready-show failed");
                }
            }

            public bool ShouldIgnoreWorkbookActivateDuringCaseProtection(Excel.Workbook workbook)
            {
                return false;
            }

            public bool ShouldIgnoreTaskPaneRefreshDuringCaseProtection(string reason, Excel.Workbook workbook, Excel.Window window)
            {
                return false;
            }

            public bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
            {
                return false;
            }

            public void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason)
            {
            }
        }
    }
}
