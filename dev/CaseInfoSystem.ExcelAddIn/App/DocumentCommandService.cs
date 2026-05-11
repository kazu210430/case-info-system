using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal interface IScreenUpdatingExecutionBridge
    {
        void Execute(Action action);
    }

    internal interface ITaskPaneRefreshSuppressionBridge
    {
        IDisposable Enter(string reason);
    }

    internal interface IActiveTaskPaneRefreshBridge
    {
        void RequestRefresh(string reason);
    }

    internal interface IKernelSheetPaneRefreshBridge
    {
        bool ShowKernelSheetAndRefreshPane(WorkbookContext context, string sheetCodeName, string reason);
    }

    internal sealed class ThisAddInScreenUpdatingExecutionBridge : IScreenUpdatingExecutionBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInScreenUpdatingExecutionBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public void Execute(Action action)
        {
            _addIn.RunWithScreenUpdatingSuspended(action);
        }
    }

    internal sealed class ThisAddInTaskPaneRefreshSuppressionBridge : ITaskPaneRefreshSuppressionBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInTaskPaneRefreshSuppressionBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public IDisposable Enter(string reason)
        {
            return _addIn.SuppressTaskPaneRefresh(reason);
        }
    }

    internal sealed class ThisAddInActiveTaskPaneRefreshBridge : IActiveTaskPaneRefreshBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInActiveTaskPaneRefreshBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public void RequestRefresh(string reason)
        {
            _addIn.RefreshActiveTaskPane(reason);
        }
    }

    internal sealed class ThisAddInKernelSheetPaneRefreshBridge : IKernelSheetPaneRefreshBridge
    {
        private readonly ThisAddIn _addIn;

        internal ThisAddInKernelSheetPaneRefreshBridge(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public bool ShowKernelSheetAndRefreshPane(WorkbookContext context, string sheetCodeName, string reason)
        {
            Excel.Workbook displayedWorkbook;
            return _addIn.ShowKernelSheetAndRefreshPaneFromHome(context, sheetCodeName, reason, out displayedWorkbook);
        }
    }

    internal sealed class DocumentCommandService
    {
        private const string DocumentActionKind = "doc";
        private const string CaseListActionKind = "caselist";

        private readonly DocumentExecutionModeService _documentExecutionModeService;
        private readonly DocumentExecutionEligibilityService _documentExecutionEligibilityService;
        private readonly DocumentCreateService _documentCreateService;
        private readonly AccountingSetCommandService _accountingSetCommandService;
        private readonly CaseListRegistrationService _caseListRegistrationService;
        private readonly CaseContextFactory _caseContextFactory;
        private readonly ExcelInteropService _excelInteropService;
        private readonly IScreenUpdatingExecutionBridge _screenUpdatingExecutionBridge;
        private readonly ITaskPaneRefreshSuppressionBridge _taskPaneRefreshSuppressionBridge;
        private readonly IKernelSheetPaneRefreshBridge _kernelSheetPaneRefreshBridge;
        private readonly Logger _logger;

        internal DocumentCommandService(
            ThisAddIn addIn,
            IScreenUpdatingExecutionBridge screenUpdatingExecutionBridge,
            ITaskPaneRefreshSuppressionBridge taskPaneRefreshSuppressionBridge,
            IActiveTaskPaneRefreshBridge activeTaskPaneRefreshBridge,
            DocumentExecutionModeService documentExecutionModeService,
            DocumentExecutionEligibilityService documentExecutionEligibilityService,
            DocumentCreateService documentCreateService,
            AccountingSetCommandService accountingSetCommandService,
            CaseListRegistrationService caseListRegistrationService,
            CaseContextFactory caseContextFactory,
            ExcelInteropService excelInteropService,
            Logger logger)
            : this(
                screenUpdatingExecutionBridge,
                taskPaneRefreshSuppressionBridge,
                activeTaskPaneRefreshBridge,
                new ThisAddInKernelSheetPaneRefreshBridge(addIn),
                documentExecutionModeService,
                documentExecutionEligibilityService,
                documentCreateService,
                accountingSetCommandService,
                caseListRegistrationService,
                caseContextFactory,
                excelInteropService,
                logger)
        {
        }

        internal DocumentCommandService(
            IScreenUpdatingExecutionBridge screenUpdatingExecutionBridge,
            ITaskPaneRefreshSuppressionBridge taskPaneRefreshSuppressionBridge,
            IActiveTaskPaneRefreshBridge activeTaskPaneRefreshBridge,
            IKernelSheetPaneRefreshBridge kernelSheetPaneRefreshBridge,
            DocumentExecutionModeService documentExecutionModeService,
            DocumentExecutionEligibilityService documentExecutionEligibilityService,
            DocumentCreateService documentCreateService,
            AccountingSetCommandService accountingSetCommandService,
            CaseListRegistrationService caseListRegistrationService,
            CaseContextFactory caseContextFactory,
            ExcelInteropService excelInteropService,
            Logger logger)
        {
            _screenUpdatingExecutionBridge = screenUpdatingExecutionBridge ?? throw new ArgumentNullException(nameof(screenUpdatingExecutionBridge));
            _taskPaneRefreshSuppressionBridge = taskPaneRefreshSuppressionBridge ?? throw new ArgumentNullException(nameof(taskPaneRefreshSuppressionBridge));
            if (activeTaskPaneRefreshBridge == null)
            {
                throw new ArgumentNullException(nameof(activeTaskPaneRefreshBridge));
            }

            _kernelSheetPaneRefreshBridge = kernelSheetPaneRefreshBridge ?? throw new ArgumentNullException(nameof(kernelSheetPaneRefreshBridge));
            _documentExecutionModeService = documentExecutionModeService ?? throw new ArgumentNullException(nameof(documentExecutionModeService));
            _documentExecutionEligibilityService = documentExecutionEligibilityService ?? throw new ArgumentNullException(nameof(documentExecutionEligibilityService));
            _documentCreateService = documentCreateService ?? throw new ArgumentNullException(nameof(documentCreateService));
            _accountingSetCommandService = accountingSetCommandService ?? throw new ArgumentNullException(nameof(accountingSetCommandService));
            _caseListRegistrationService = caseListRegistrationService ?? throw new ArgumentNullException(nameof(caseListRegistrationService));
            _caseContextFactory = caseContextFactory ?? throw new ArgumentNullException(nameof(caseContextFactory));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal void Execute(Excel.Workbook workbook, string actionKind, string key)
        {
            if (workbook == null)
            {
                throw new InvalidOperationException("CASE workbook was not found.");
            }

            DocumentCommandActionRoute route = DocumentCommandActionRoutePolicy.Decide(actionKind);

            if (route == DocumentCommandActionRoute.Document)
            {
                ExecuteDocumentAction(workbook, key);
                return;
            }

            if (route == DocumentCommandActionRoute.Accounting)
            {
                _accountingSetCommandService.Execute(workbook);
                _logger.Info("Accounting action was handled by VSTO flow.");
                return;
            }

            if (route == DocumentCommandActionRoute.Unsupported)
            {
                _logger.Warn("Unsupported task pane action was blocked. actionKind=" + (actionKind ?? string.Empty) + ", key=" + (key ?? string.Empty));
                throw new InvalidOperationException("未対応の操作です。actionKind=" + (actionKind ?? string.Empty));
            }

            CaseListRegistrationResult completedRegistrationResult = null;
            WorkbookContext kernelSheetTransitionContext = null;
            bool shouldShowKernelCaseList = false;
            _screenUpdatingExecutionBridge.Execute(() =>
            {
                CaseListRegistrationResult registrationResult = _caseListRegistrationService.Execute(workbook);
                if (!registrationResult.Success)
                {
                    throw new InvalidOperationException(registrationResult.Message);
                }

                var context = _caseContextFactory.CreateForCaseListRegistration(workbook);
                if (context == null)
                {
                    _logger.Info("Case list row normalization skipped because context could not be resolved.");
                    string saveFailureMessageForWorkbook;
                    if (!TrySaveKernelWorkbook(workbook, registrationResult, out saveFailureMessageForWorkbook))
                    {
                        throw new InvalidOperationException(saveFailureMessageForWorkbook);
                    }

                    return;
                }

                bool normalized = _excelInteropService.TryNormalizeCaseListRowHeight(context);
                if (!normalized)
                {
                    _logger.Info("Case list row normalization was skipped or failed.");
                }

                string saveFailureMessageForKernel;
                if (!TrySaveKernelWorkbook(context.KernelWorkbook, registrationResult, out saveFailureMessageForKernel))
                {
                    throw new InvalidOperationException(saveFailureMessageForKernel);
                }

                kernelSheetTransitionContext = CreateKernelSheetTransitionContext(context);
                completedRegistrationResult = registrationResult;
                shouldShowKernelCaseList = true;
            });

            // Keep the already-open CASE pane as-is; the Kernel sheet transition owns this route.
            if (shouldShowKernelCaseList)
            {
                string kernelTransitionSheetCodeName = "shCaseList";
                string kernelTransitionReason = "DocumentCommandService.Execute";
                bool paneRefreshed = _kernelSheetPaneRefreshBridge.ShowKernelSheetAndRefreshPane(
                    kernelSheetTransitionContext,
                    kernelTransitionSheetCodeName,
                    kernelTransitionReason);
                if (!paneRefreshed)
                {
                    _logger.Info("Kernel case list pane refresh by unified add-in was not available.");
                }
            }

            if (completedRegistrationResult != null && completedRegistrationResult.Success)
            {
                ShowCaseListRegistrationMessage(completedRegistrationResult);
            }
            return;
        }

        /// <summary>
        /// メソッド: 案件一覧登録完了メッセージを表示する。
        /// 引数: registrationResult - 案件一覧登録結果。
        /// 戻り値: なし。
        /// 副作用: メッセージダイアログを表示する。
        /// </summary>
        private void ShowCaseListRegistrationMessage(CaseListRegistrationResult registrationResult)
        {
            if (registrationResult == null)
            {
                throw new ArgumentNullException(nameof(registrationResult));
            }

            string message = registrationResult.Message ?? "案件一覧登録が完了しました。";
            Excel.Window activeWindowBeforeMessage = _excelInteropService.GetActiveWindow();
            ExcelWindowOwner owner = ExcelWindowOwner.From(activeWindowBeforeMessage);
            try
            {
                CompletionNoticeForm.ShowNotice(owner, "案件情報System", message);
            }
            finally
            {
                if (owner != null)
                {
                    owner.Dispose();
                }
            }
        }

        private bool TrySaveKernelWorkbook(Excel.Workbook kernelWorkbook, CaseListRegistrationResult registrationResult, out string failureMessage)
        {
            failureMessage = string.Empty;
            if (kernelWorkbook == null)
            {
                return true;
            }

            try
            {
                kernelWorkbook.Save();
                _logger.Info("Kernel workbook saved after case-list registration. row=" + registrationResult.RegisteredRow.ToString());
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel workbook save after case-list registration failed.", ex);
                failureMessage = "案件一覧登録後の保存に失敗しました。Excel 上で保存状態を確認してください。詳細: " + ex.Message;
                return false;
            }
        }

        private WorkbookContext CreateKernelSheetTransitionContext(CaseContext context)
        {
            Excel.Workbook kernelWorkbook = context == null ? null : context.KernelWorkbook;
            if (kernelWorkbook == null)
            {
                return null;
            }

            Excel.Worksheet caseListWorksheet = context.CaseListWorksheet;
            return new WorkbookContext(
                kernelWorkbook,
                null,
                WorkbookRole.Kernel,
                context.SystemRoot,
                _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(kernelWorkbook),
                caseListWorksheet == null ? string.Empty : (caseListWorksheet.CodeName ?? string.Empty));
        }

        /// <summary>
        /// メソッド: 文書作成ボタン押下を VSTO 本体で実行する。
        /// 引数: workbook - 対象 CASE ブック, key - 文書キー。
        /// 戻り値: なし。
        /// 副作用: Word 文書を生成し、未許可条件では例外を送出する。
        /// </summary>
        private void ExecuteDocumentAction(Excel.Workbook workbook, string key)
        {
            Stopwatch totalStopwatch = Stopwatch.StartNew();
            Stopwatch phaseStopwatch = Stopwatch.StartNew();
            using (_taskPaneRefreshSuppressionBridge.Enter("DocumentCommandService.ExecuteDocumentAction"))
            {
                _screenUpdatingExecutionBridge.Execute(() =>
                {
                    DocumentExecutionMode executionMode = _documentExecutionModeService.GetConfiguredMode();
                    DocumentExecutionEligibility eligibility = _documentExecutionEligibilityService.Evaluate(workbook, DocumentActionKind, key);
                    _logger.Debug(
                        "ExecuteDocumentAction",
                        "EligibilityEvaluated elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed)
                        + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed)
                        + " key=" + (key ?? string.Empty)
                        + " canExecute=" + eligibility.CanExecuteInVsto.ToString());
                    phaseStopwatch.Restart();
                    DocumentCommandPreconditionDecision preconditionDecision = DocumentCommandPreconditionPolicy.Decide(
                        eligibility.CanExecuteInVsto);
                    DocumentCommandExecutionDecision executionDecision = DocumentCommandExecutionDecisionPolicy.Decide(preconditionDecision);
                    if (executionDecision == DocumentCommandExecutionDecision.ThrowBecauseIneligible)
                    {
                        _logger.Warn(
                            "Document action was blocked because the template was not eligible for VSTO execution."
                            + " mode=" + executionMode.ToString()
                            + ", key=" + (key ?? string.Empty)
                            + ", reason=" + (eligibility.Reason ?? string.Empty));
                        throw new InvalidOperationException("文書作成を実行できませんでした。理由: " + (eligibility.Reason ?? "eligible 判定に失敗しました。"));
                    }

                    phaseStopwatch.Restart();
                    _documentCreateService.Execute(workbook, eligibility.TemplateSpec, eligibility.CaseContext);
                    _logger.Debug(
                        "ExecuteDocumentAction",
                        "DocumentCreateCompleted elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed)
                        + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed)
                        + " key=" + (key ?? string.Empty));
                    _logger.Info(
                        "Document action was handled by VSTO document flow."
                        + " mode=" + executionMode.ToString()
                        + ", key=" + (key ?? string.Empty)
                        + ", templateFile=" + (eligibility.TemplateSpec == null ? string.Empty : (eligibility.TemplateSpec.TemplateFileName ?? string.Empty)));
                });
            }
        }

        private static string FormatElapsedSeconds(TimeSpan elapsed)
        {
            return elapsed.TotalSeconds.ToString("0.000");
        }
    }
}
