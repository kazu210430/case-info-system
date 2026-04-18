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
    public class DocumentCommandServiceTests
    {
        [Fact]
        public void Execute_WhenKernelSaveFails_DoesNotAdvanceToSuccessFlow()
        {
            var logs = new List<string>();
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            int refreshCalls = 0;
            int showKernelSheetCalls = 0;
            int showNoticeCalls = 0;
            var kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\kernel.xlsx",
                SaveBehavior = () => throw new InvalidOperationException("save failed")
            };
            var registrationResult = new CaseListRegistrationResult
            {
                Success = true,
                RegisteredRow = 12,
                Message = "ok"
            };

            CompletionNoticeForm.OnShowNotice = (owner, title, message) => showNoticeCalls++;
            addIn.ShowKernelSheetAndRefreshPaneHandler = (sheetCodeName, reason) =>
            {
                showKernelSheetCalls++;
                return true;
            };

            var service = new DocumentCommandService(
                addIn,
                new InlineScreenUpdatingExecutionBridge(),
                new NoOpTaskPaneRefreshSuppressionBridge(),
                new CollectingActiveTaskPaneRefreshBridge(reason => refreshCalls++),
                new DocumentExecutionModeService(OrchestrationTestSupport.CreateLogger(new List<string>()), new ExcelInteropService()),
                new DocumentExecutionEligibilityService(),
                new DocumentExecutionPolicyService(),
                new DocumentCreateService(),
                new AccountingSetCommandService(),
                new CaseListRegistrationService
                {
                    OnExecute = workbook => registrationResult
                },
                new CaseContextFactory
                {
                    OnCreateForCaseListRegistration = workbook => new CaseContext
                    {
                        KernelWorkbook = kernelWorkbook
                    }
                },
                new ExcelInteropService
                {
                    OnTryNormalizeCaseListRowHeight = context => true
                },
                OrchestrationTestSupport.CreateLogger(logs));

            var exception = Assert.Throws<InvalidOperationException>(() => service.Execute(new Excel.Workbook(), "caselist", "ignored"));

            Assert.Contains("保存に失敗しました", exception.Message);
            Assert.Equal(1, kernelWorkbook.SaveCallCount);
            Assert.Equal(0, refreshCalls);
            Assert.Equal(0, showKernelSheetCalls);
            Assert.Equal(0, showNoticeCalls);
            Assert.Contains(logs, message => message.Contains("Kernel workbook save after case-list registration failed."));
        }

        [Fact]
        public void Execute_WhenKernelSaveSucceeds_CompletesSuccessFlow()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            int refreshCalls = 0;
            int showKernelSheetCalls = 0;
            int showNoticeCalls = 0;
            string shownMessage = string.Empty;
            var kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\kernel.xlsx"
            };
            var registrationResult = new CaseListRegistrationResult
            {
                Success = true,
                RegisteredRow = 8,
                Message = "案件一覧登録が完了しました。"
            };

            CompletionNoticeForm.OnShowNotice = (owner, title, message) =>
            {
                showNoticeCalls++;
                shownMessage = message;
            };
            addIn.ShowKernelSheetAndRefreshPaneHandler = (sheetCodeName, reason) =>
            {
                showKernelSheetCalls++;
                return true;
            };

            var service = new DocumentCommandService(
                addIn,
                new InlineScreenUpdatingExecutionBridge(),
                new NoOpTaskPaneRefreshSuppressionBridge(),
                new CollectingActiveTaskPaneRefreshBridge(reason => refreshCalls++),
                new DocumentExecutionModeService(OrchestrationTestSupport.CreateLogger(new List<string>()), new ExcelInteropService()),
                new DocumentExecutionEligibilityService(),
                new DocumentExecutionPolicyService(),
                new DocumentCreateService(),
                new AccountingSetCommandService(),
                new CaseListRegistrationService
                {
                    OnExecute = workbook => registrationResult
                },
                new CaseContextFactory
                {
                    OnCreateForCaseListRegistration = workbook => new CaseContext
                    {
                        KernelWorkbook = kernelWorkbook
                    }
                },
                new ExcelInteropService
                {
                    OnTryNormalizeCaseListRowHeight = context => true
                },
                OrchestrationTestSupport.CreateLogger(new List<string>()));

            service.Execute(new Excel.Workbook(), "caselist", "ignored");

            Assert.Equal(1, kernelWorkbook.SaveCallCount);
            Assert.True(kernelWorkbook.Saved);
            Assert.Equal(1, refreshCalls);
            Assert.Equal(1, showKernelSheetCalls);
            Assert.Equal(1, showNoticeCalls);
            Assert.Equal(registrationResult.Message, shownMessage);
        }

        private sealed class InlineScreenUpdatingExecutionBridge : IScreenUpdatingExecutionBridge
        {
            public void Execute(Action action)
            {
                action?.Invoke();
            }
        }

        private sealed class NoOpTaskPaneRefreshSuppressionBridge : ITaskPaneRefreshSuppressionBridge
        {
            public IDisposable Enter(string reason)
            {
                return new NoOpDisposable();
            }
        }

        private sealed class CollectingActiveTaskPaneRefreshBridge : IActiveTaskPaneRefreshBridge
        {
            private readonly Action<string> _onRequest;

            public CollectingActiveTaskPaneRefreshBridge(Action<string> onRequest)
            {
                _onRequest = onRequest;
            }

            public void RequestRefresh(string reason)
            {
                _onRequest?.Invoke(reason);
            }
        }

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
