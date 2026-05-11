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
                KernelWorkbook = kernelWorkbook,
                Message = "ok"
            };

            CompletionNoticeForm.OnShowNotice = (owner, title, message) => showNoticeCalls++;
            addIn.ShowKernelSheetAndRefreshPaneFromHomeHandler = (context, sheetCodeName, reason) =>
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
                        KernelWorkbook = kernelWorkbook,
                        SystemRoot = @"C:\root"
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
            WorkbookContext shownContext = null;
            var kernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\kernel.xlsx"
            };
            var registrationResult = new CaseListRegistrationResult
            {
                Success = true,
                RegisteredRow = 8,
                KernelWorkbook = kernelWorkbook,
                Message = "案件一覧登録が完了しました。"
            };

            CompletionNoticeForm.OnShowNotice = (owner, title, message) =>
            {
                showNoticeCalls++;
                shownMessage = message;
            };
            addIn.ShowKernelSheetAndRefreshPaneFromHomeHandler = (context, sheetCodeName, reason) =>
            {
                shownContext = context;
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
                        KernelWorkbook = kernelWorkbook,
                        CaseListWorksheet = new Excel.Worksheet
                        {
                            CodeName = "shCaseList"
                        },
                        SystemRoot = @"C:\root"
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
            Assert.Equal(0, refreshCalls);
            Assert.Equal(1, showKernelSheetCalls);
            Assert.Equal(1, showNoticeCalls);
            Assert.Equal(registrationResult.Message, shownMessage);
            Assert.NotNull(shownContext);
            Assert.Same(kernelWorkbook, shownContext.Workbook);
            Assert.Equal(WorkbookRole.Kernel, shownContext.Role);
            Assert.Equal(@"C:\root", shownContext.SystemRoot);
            Assert.Equal(kernelWorkbook.FullName, shownContext.WorkbookFullName);
            Assert.Equal("shCaseList", shownContext.ActiveSheetCodeName);
        }

        [Fact]
        public void Execute_WhenRegistrationResultHasKernelWorkbook_SavesRegisteredKernelWorkbook()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            int refreshCalls = 0;
            int showKernelSheetCalls = 0;
            int showNoticeCalls = 0;
            WorkbookContext shownContext = null;
            var registeredKernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\registered-kernel.xlsx"
            };
            var contextKernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\context-kernel.xlsx"
            };
            var registrationResult = new CaseListRegistrationResult
            {
                Success = true,
                RegisteredRow = 21,
                KernelWorkbook = registeredKernelWorkbook,
                Message = "registered"
            };

            CompletionNoticeForm.OnShowNotice = (owner, title, message) => showNoticeCalls++;
            addIn.ShowKernelSheetAndRefreshPaneFromHomeHandler = (context, sheetCodeName, reason) =>
            {
                shownContext = context;
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
                        KernelWorkbook = contextKernelWorkbook,
                        CaseListWorksheet = new Excel.Worksheet
                        {
                            CodeName = "shCaseList"
                        },
                        SystemRoot = @"C:\root"
                    }
                },
                new ExcelInteropService
                {
                    OnTryNormalizeCaseListRowHeight = context => true
                },
                OrchestrationTestSupport.CreateLogger(new List<string>()));

            service.Execute(new Excel.Workbook(), "caselist", "ignored");

            Assert.Equal(1, registeredKernelWorkbook.SaveCallCount);
            Assert.True(registeredKernelWorkbook.Saved);
            Assert.Equal(0, contextKernelWorkbook.SaveCallCount);
            Assert.Equal(0, refreshCalls);
            Assert.Equal(1, showKernelSheetCalls);
            Assert.Equal(1, showNoticeCalls);
            Assert.NotNull(shownContext);
            Assert.Same(contextKernelWorkbook, shownContext.Workbook);
        }

        [Fact]
        public void Execute_WhenCaseListContextCannotResolve_SavesRegisteredKernelWorkbookAndDoesNotSaveCaseWorkbook()
        {
            var addIn = new CaseInfoSystem.ExcelAddIn.ThisAddIn();
            int refreshCalls = 0;
            int showKernelSheetCalls = 0;
            int showNoticeCalls = 0;
            var caseWorkbook = new Excel.Workbook
            {
                FullName = @"C:\case.xlsx"
            };
            var registeredKernelWorkbook = new Excel.Workbook
            {
                FullName = @"C:\registered-kernel.xlsx"
            };
            var registrationResult = new CaseListRegistrationResult
            {
                Success = true,
                RegisteredRow = 22,
                KernelWorkbook = registeredKernelWorkbook,
                Message = "registered"
            };

            CompletionNoticeForm.OnShowNotice = (owner, title, message) => showNoticeCalls++;
            addIn.ShowKernelSheetAndRefreshPaneFromHomeHandler = (context, sheetCodeName, reason) =>
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
                new DocumentCreateService(),
                new AccountingSetCommandService(),
                new CaseListRegistrationService
                {
                    OnExecute = workbook => registrationResult
                },
                new CaseContextFactory
                {
                    OnCreateForCaseListRegistration = workbook => null
                },
                new ExcelInteropService
                {
                    OnTryNormalizeCaseListRowHeight = context => throw new InvalidOperationException("normalization should not run without context")
                },
                OrchestrationTestSupport.CreateLogger(new List<string>()));

            service.Execute(caseWorkbook, "caselist", "ignored");

            Assert.Equal(0, caseWorkbook.SaveCallCount);
            Assert.Equal(1, registeredKernelWorkbook.SaveCallCount);
            Assert.True(registeredKernelWorkbook.Saved);
            Assert.Equal(0, refreshCalls);
            Assert.Equal(0, showKernelSheetCalls);
            Assert.Equal(0, showNoticeCalls);
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
