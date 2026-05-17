using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class AutomationSurfaceAdapter
    {
        private readonly Excel.Application _application;
        private readonly Logger _logger;
        private readonly string _productTitle;
        private HomeTransitionAdapter _homeTransitionAdapter;
        private ExcelInteropService _excelInteropService;
        private WorkbookRoleResolver _workbookRoleResolver;
        private KernelWorkbookService _kernelWorkbookService;
        private KernelCommandService _kernelCommandService;
        private KernelUserDataReflectionService _kernelUserDataReflectionService;
        private WorkbookRibbonCommandService _workbookRibbonCommandService;
        private WorkbookCaseTaskPaneRefreshCommandService _workbookCaseTaskPaneRefreshCommandService;
        private WorkbookResetCommandService _workbookResetCommandService;

        internal AutomationSurfaceAdapter(Excel.Application application, Logger logger, string productTitle)
        {
            _application = application;
            _logger = logger;
            _productTitle = productTitle ?? string.Empty;
        }

        internal void Configure(
            HomeTransitionAdapter homeTransitionAdapter,
            ExcelInteropService excelInteropService,
            WorkbookRoleResolver workbookRoleResolver,
            KernelWorkbookService kernelWorkbookService,
            KernelCommandService kernelCommandService,
            KernelUserDataReflectionService kernelUserDataReflectionService,
            WorkbookRibbonCommandService workbookRibbonCommandService,
            WorkbookCaseTaskPaneRefreshCommandService workbookCaseTaskPaneRefreshCommandService,
            WorkbookResetCommandService workbookResetCommandService)
        {
            _homeTransitionAdapter = homeTransitionAdapter;
            _excelInteropService = excelInteropService;
            _workbookRoleResolver = workbookRoleResolver;
            _kernelWorkbookService = kernelWorkbookService;
            _kernelCommandService = kernelCommandService;
            _kernelUserDataReflectionService = kernelUserDataReflectionService;
            _workbookRibbonCommandService = workbookRibbonCommandService;
            _workbookCaseTaskPaneRefreshCommandService = workbookCaseTaskPaneRefreshCommandService;
            _workbookResetCommandService = workbookResetCommandService;
        }

        internal void ShowKernelHomeFromAutomation()
        {
            _logger?.Info("Kernel home requested from COM automation.");
            if (_logger == null)
            {
                ExcelAddInTraceLogWriter.Write("Kernel home requested from COM automation.");
            }

            _homeTransitionAdapter.ShowKernelHomePlaceholderWithExternalWorkbookSuppressionForNewSession("KernelAutomationService.ShowHome");
        }

        internal void LogAutomationFailure(string message, Exception ex)
        {
            if (_logger != null)
            {
                _logger.Error(message, ex);
                return;
            }

            ExcelAddInTraceLogWriter.Write((message ?? string.Empty) + " exception=" + (ex == null ? string.Empty : ex.ToString()));
        }

        internal void ReflectKernelUserDataToAccountingSet()
        {
            _logger?.Info("Kernel user data reflection requested from COM automation. target=AccountingSet");
            WorkbookContext context = ResolveKernelReflectionContextForAutomation();
            _kernelUserDataReflectionService?.ReflectToAccountingSetOnly(context);
        }

        internal void ReflectKernelUserDataToBaseHome()
        {
            _logger?.Info("Kernel user data reflection requested from COM automation. target=BaseHome");
            WorkbookContext context = ResolveKernelReflectionContextForAutomation();
            _kernelUserDataReflectionService?.ReflectToBaseHomeOnly(context);
        }

        internal void ShowActiveWorkbookCustomDocumentProperties()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            _workbookRibbonCommandService?.ShowCustomDocumentProperties(targetWorkbook);
        }

        internal void SelectAndSaveActiveWorkbookSystemRoot()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            _workbookRibbonCommandService?.SelectAndSaveSystemRoot(targetWorkbook);
        }

        internal void RefreshActiveWorkbookCaseTaskPane()
        {
            Excel.Workbook workbook = ResolveRibbonTargetWorkbook();
            if (_workbookCaseTaskPaneRefreshCommandService == null)
            {
                MessageBox.Show("Pane 更新サービスを利用できません。", _productTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _workbookCaseTaskPaneRefreshCommandService.Refresh(workbook);
        }

        internal void CopySampleColumnBToHome()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            _workbookRibbonCommandService?.CopySampleColumnBToHome(targetWorkbook);
        }

        internal void UpdateBaseDefinitionFromRibbon()
        {
            if (_kernelCommandService == null)
            {
                MessageBox.Show("Base定義更新サービスを利用できません。", _productTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                WorkbookContext context = ResolveKernelCommandContextForRibbon();
                _kernelCommandService.Execute(context, KernelNavigationActionIds.SyncBaseHomeFieldInventory);
            }
            catch (Exception ex)
            {
                _logger?.Error("UpdateBaseDefinitionFromRibbon failed.", ex);
                MessageBox.Show("Base定義更新を実行できませんでした。ログを確認してください。", _productTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        internal void ResetActiveWorkbookForDistribution()
        {
            Excel.Workbook targetWorkbook = ResolveRibbonTargetWorkbook();
            WorkbookResetResult result = _workbookResetCommandService == null
                ? new WorkbookResetResult
                {
                    Success = false,
                    Message = "配布前リセットサービスを利用できません。"
                }
                : _workbookResetCommandService.Execute(targetWorkbook);
            _workbookResetCommandService?.ShowResult(result);
        }

        private Excel.Workbook ResolveRibbonTargetWorkbook()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            if (activeWorkbook != null)
            {
                return activeWorkbook;
            }

            if (_excelInteropService == null)
            {
                return null;
            }

            var openWorkbooks = _excelInteropService.GetOpenWorkbooks();
            return openWorkbooks.Count == 1 ? openWorkbooks[0] : null;
        }

        private WorkbookContext ResolveKernelCommandContextForRibbon()
        {
            Excel.Workbook workbook = ResolveRibbonTargetWorkbook();
            string systemRoot = string.Empty;
            if (workbook != null && _excelInteropService != null)
            {
                systemRoot = _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT");
            }

            if ((workbook == null || string.IsNullOrWhiteSpace(systemRoot)) && _kernelWorkbookService != null)
            {
                string boundSystemRoot;
                Excel.Workbook boundWorkbook;
                if (_kernelWorkbookService.TryGetValidHomeWorkbookBinding(out boundWorkbook, out boundSystemRoot))
                {
                    workbook = boundWorkbook;
                    systemRoot = boundSystemRoot;
                }
            }

            if (workbook == null || _excelInteropService == null)
            {
                throw new InvalidOperationException("Kernel workbook context was not available for Base definition update.");
            }

            WorkbookRole role = _workbookRoleResolver == null
                ? WorkbookRole.Unknown
                : _workbookRoleResolver.Resolve(workbook);
            return new WorkbookContext(
                workbook,
                TryGetActiveWindow(),
                role,
                systemRoot,
                _excelInteropService.GetWorkbookFullName(workbook),
                _excelInteropService.GetActiveSheetCodeName(workbook));
        }

        private WorkbookContext ResolveKernelReflectionContextForAutomation()
        {
            Excel.Workbook workbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            string systemRoot = _excelInteropService == null || workbook == null
                ? string.Empty
                : _excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT");

            if (workbook == null && _kernelWorkbookService != null)
            {
                string boundSystemRoot;
                if (_kernelWorkbookService.TryGetValidHomeWorkbookBinding(out workbook, out boundSystemRoot))
                {
                    systemRoot = boundSystemRoot;
                }
            }

            if (workbook == null || _excelInteropService == null || _workbookRoleResolver == null)
            {
                throw new InvalidOperationException("Kernel workbook context was not available for user-data reflection.");
            }

            return new WorkbookContext(
                workbook,
                _excelInteropService.GetActiveWindow(),
                _workbookRoleResolver.Resolve(workbook),
                systemRoot,
                _excelInteropService.GetWorkbookFullName(workbook),
                _excelInteropService.GetActiveSheetCodeName(workbook));
        }

        private Excel.Window TryGetActiveWindow()
        {
            try
            {
                return _application == null ? null : _application.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }
    }
}
