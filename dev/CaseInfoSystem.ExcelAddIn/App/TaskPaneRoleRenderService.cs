using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRoleRenderService
    {
        private readonly CasePaneSnapshotRenderService _casePaneSnapshotRenderService;
        private readonly CasePaneCacheRefreshNotificationService _casePaneCacheRefreshNotificationService;
        private readonly Logger _logger;

        internal TaskPaneRoleRenderService(
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            CasePaneCacheRefreshNotificationService casePaneCacheRefreshNotificationService,
            Logger logger)
        {
            _casePaneSnapshotRenderService = casePaneSnapshotRenderService ?? throw new ArgumentNullException(nameof(casePaneSnapshotRenderService));
            _casePaneCacheRefreshNotificationService = casePaneCacheRefreshNotificationService ?? throw new ArgumentNullException(nameof(casePaneCacheRefreshNotificationService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal void Render(TaskPaneHost host, WorkbookContext context, string reason)
        {
            if (host == null)
            {
                throw new ArgumentNullException(nameof(host));
            }

            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (host.Control is KernelNavigationControl kernelControl)
            {
                RenderKernelHost(kernelControl, context);
                return;
            }

            if (host.Control is AccountingNavigationControl accountingControl)
            {
                RenderAccountingHost(accountingControl, context);
                return;
            }

            if (host.Control is DocumentButtonsControl caseControl)
            {
                RenderCaseHost(caseControl, context, reason);
            }
        }

        private void RenderKernelHost(KernelNavigationControl kernelControl, WorkbookContext context)
        {
            _logger.Info("RenderHost start. role=Kernel, workbook=" + (context.WorkbookFullName ?? string.Empty));
            kernelControl.Render(KernelNavigationDefinitions.CreateForSheet(context.ActiveSheetCodeName));
            _logger.Info("RenderHost completed. role=Kernel, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        private void RenderAccountingHost(AccountingNavigationControl accountingControl, WorkbookContext context)
        {
            _logger.Info("RenderHost start. role=Accounting, workbook=" + (context.WorkbookFullName ?? string.Empty));
            accountingControl.Render(AccountingNavigationDefinitions.CreateForSheet(context.ActiveSheetCodeName));
            _logger.Info("RenderHost completed. role=Accounting, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        private void RenderCaseHost(DocumentButtonsControl caseControl, WorkbookContext context, string reason)
        {
            _logger.Info("RenderHost start. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
            bool? originalWorkbookSavedState = _casePaneCacheRefreshNotificationService.TryGetWorkbookSavedState(context.Workbook);
            CasePaneSnapshotRenderService.CasePaneSnapshotRenderResult renderResult = _casePaneSnapshotRenderService.Render(caseControl, context.Workbook);
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult = renderResult.BuildResult;
            string snapshotText = buildResult.SnapshotText;
            _logger.Info("RenderHost snapshot acquired. role=Case, length=" + snapshotText.Length.ToString());
            TaskPaneSnapshot snapshot = renderResult.Snapshot;
            _logger.Info("RenderHost snapshot parsed. role=Case, hasError=" + snapshot.HasError.ToString() + ", tabs=" + snapshot.Tabs.Count.ToString() + ", docs=" + snapshot.DocButtons.Count.ToString());
            _casePaneCacheRefreshNotificationService.NotifyCasePaneUpdatedIfNeeded(context.Workbook, reason, buildResult, originalWorkbookSavedState);
            _logger.Info("RenderHost completed. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }
    }
}
