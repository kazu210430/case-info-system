using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneNonCaseActionHandler
    {
        private readonly ExcelInteropService _excelInteropService;
        private readonly KernelCommandService _kernelCommandService;
        private readonly AccountingSheetCommandService _accountingSheetCommandService;
        private readonly AccountingInternalCommandService _accountingInternalCommandService;
        private readonly UserErrorService _userErrorService;
        private readonly Logger _logger;
        private readonly Func<string, TaskPaneHost> _resolveHost;
        private readonly Action<TaskPaneHost, WorkbookContext, string> _renderHost;
        private readonly Func<TaskPaneHost, string, bool> _tryShowHost;

        internal TaskPaneNonCaseActionHandler(
            ExcelInteropService excelInteropService,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            AccountingInternalCommandService accountingInternalCommandService,
            UserErrorService userErrorService,
            Logger logger,
            Func<string, TaskPaneHost> resolveHost,
            Action<TaskPaneHost, WorkbookContext, string> renderHost,
            Func<TaskPaneHost, string, bool> tryShowHost)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _kernelCommandService = kernelCommandService ?? throw new ArgumentNullException(nameof(kernelCommandService));
            _accountingSheetCommandService = accountingSheetCommandService ?? throw new ArgumentNullException(nameof(accountingSheetCommandService));
            _accountingInternalCommandService = accountingInternalCommandService ?? throw new ArgumentNullException(nameof(accountingInternalCommandService));
            _userErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _resolveHost = resolveHost ?? throw new ArgumentNullException(nameof(resolveHost));
            _renderHost = renderHost ?? throw new ArgumentNullException(nameof(renderHost));
            _tryShowHost = tryShowHost ?? throw new ArgumentNullException(nameof(tryShowHost));
        }

        internal void HandleKernelActionInvoked(string windowKey, KernelNavigationActionEventArgs e)
        {
            const string actionName = "KernelControl_ActionInvoked";
            if (!TryResolveActionTarget(windowKey, actionName, out TaskPaneHost host, out Excel.Workbook workbook))
            {
                return;
            }

            WorkbookContext context = BuildWorkbookContext(
                host,
                workbook,
                WorkbookRole.Kernel,
                _excelInteropService.GetActiveSheetCodeName(workbook));
            _kernelCommandService.Execute(context, e.ActionId);
        }

        internal void HandleAccountingActionInvoked(string windowKey, AccountingNavigationActionEventArgs e)
        {
            const string actionName = "AccountingControl_ActionInvoked";
            if (!TryResolveActionTarget(windowKey, actionName, out TaskPaneHost host, out Excel.Workbook workbook))
            {
                return;
            }

            try
            {
                WorkbookContext context = BuildWorkbookContext(
                    host,
                    workbook,
                    WorkbookRole.Accounting,
                    TryGetWorksheetCodeName(workbook));
                _accountingSheetCommandService.Execute(context, e.ActionId);
                _accountingInternalCommandService.Execute(context, e.ActionId);

                WorkbookContext refreshedContext = BuildWorkbookContext(
                    host,
                    workbook,
                    WorkbookRole.Accounting,
                    _excelInteropService.GetActiveSheetCodeName(workbook));
                _renderHost(host, refreshedContext, actionName);
                _tryShowHost(host, actionName);
            }
            catch (Exception ex)
            {
                _logger.Error(actionName + " failed.", ex);
                _userErrorService.ShowUserError(actionName, ex);
            }
        }

        private bool TryResolveActionTarget(string windowKey, string actionName, out TaskPaneHost host, out Excel.Workbook workbook)
        {
            host = null;
            workbook = null;

            if (string.IsNullOrWhiteSpace(windowKey))
            {
                _logger.Warn(actionName + " skipped because windowKey was empty.");
                return false;
            }

            host = _resolveHost(windowKey);
            if (host == null)
            {
                _logger.Warn(actionName + " skipped because host was not found. windowKey=" + windowKey);
                return false;
            }

            workbook = _excelInteropService.FindOpenWorkbook(host.WorkbookFullName);
            if (workbook == null)
            {
                _logger.Warn(actionName + " skipped because workbook was not found. windowKey=" + windowKey);
                return false;
            }

            return true;
        }

        private WorkbookContext BuildWorkbookContext(TaskPaneHost host, Excel.Workbook workbook, WorkbookRole role, string activeSheetCodeName)
        {
            return new WorkbookContext(
                workbook,
                host.Window,
                role,
                _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT"),
                _excelInteropService.GetWorkbookFullName(workbook),
                activeSheetCodeName);
        }

        private static string TryGetWorksheetCodeName(Excel.Workbook workbook)
        {
            try
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;
                return worksheet == null ? string.Empty : worksheet.CodeName ?? string.Empty;
            }
            catch
            {
                // Fall back to an empty CodeName if the active sheet is unavailable during accounting pane refresh.
                return string.Empty;
            }
        }
    }
}
