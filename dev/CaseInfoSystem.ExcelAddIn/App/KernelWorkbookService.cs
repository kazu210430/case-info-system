using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookService
    {
        private readonly KernelWorkbookBindingService _bindingService;
        private readonly KernelWorkbookDisplayService _displayService;
        private readonly KernelWorkbookCloseService _closeService;

        internal KernelWorkbookService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger)
            : this(
                CreateFacadeServices(
                    application,
                    excelInteropService,
                    excelWindowRecoveryService,
                    kernelCaseInteractionState,
                    logger))
        {
        }

        internal KernelWorkbookService(
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger,
            KernelWorkbookServiceTestHooks testHooks)
            : this(
                CreateFacadeServices(
                    kernelCaseInteractionState,
                    logger,
                    testHooks))
        {
        }

        internal KernelWorkbookService(
            KernelWorkbookBindingService bindingService,
            KernelWorkbookDisplayService displayService,
            KernelWorkbookCloseService closeService)
        {
            _bindingService = bindingService ?? throw new ArgumentNullException(nameof(bindingService));
            _displayService = displayService ?? throw new ArgumentNullException(nameof(displayService));
            _closeService = closeService ?? throw new ArgumentNullException(nameof(closeService));
        }

        internal void SetLifecycleService(KernelWorkbookLifecycleService kernelWorkbookLifecycleService)
        {
            _closeService.SetLifecycleService(kernelWorkbookLifecycleService);
        }

        internal void RegisterHomeSessionCloseObserver(Action onReadyToCloseForm, Action onFailed)
        {
            _closeService.RegisterHomeSessionCloseObserver(onReadyToCloseForm, onFailed);
        }

        internal KernelHomeSessionCloseRequestStatus RequestCloseHomeSessionFromForm(bool saveKernelWorkbook, string entryPoint)
        {
            return _closeService.RequestCloseHomeSessionFromForm(saveKernelWorkbook, entryPoint);
        }

        internal void FinalizePendingHomeSessionCloseAfterFormClosed()
        {
            _closeService.FinalizePendingHomeSessionCloseAfterFormClosed();
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return _bindingService.IsKernelWorkbook(workbook);
        }

        internal Excel.Workbook ResolveKernelWorkbook(WorkbookContext context)
        {
            return _bindingService.ResolveKernelWorkbook(context);
        }

        internal Excel.Workbook ResolveKernelWorkbook(string systemRoot)
        {
            return _bindingService.ResolveKernelWorkbook(systemRoot);
        }

        internal bool BindHomeWorkbook(WorkbookContext context)
        {
            return _bindingService.BindHomeWorkbook(context);
        }

        internal bool BindHomeWorkbook(Excel.Workbook workbook)
        {
            return _bindingService.BindHomeWorkbook(workbook);
        }

        internal void ClearHomeWorkbookBinding(string reason)
        {
            _bindingService.ClearHomeWorkbookBinding(reason);
        }

        internal bool HasValidHomeWorkbookBinding()
        {
            return _bindingService.HasValidHomeWorkbookBinding();
        }

        internal bool TryGetValidHomeWorkbookBinding(out Excel.Workbook workbook, out string systemRoot)
        {
            return _bindingService.TryGetValidHomeWorkbookBinding(out workbook, out systemRoot);
        }

        internal bool TryShowSheetByCodeName(WorkbookContext context, string sheetCodeName, string reason)
        {
            return _displayService.TryShowSheetByCodeName(context, sheetCodeName, reason);
        }

        internal bool ShouldShowHomeOnStartup(Excel.Workbook startupWorkbook = null)
        {
            return _bindingService.ShouldShowHomeOnStartup(startupWorkbook);
        }

        internal string DescribeStartupState()
        {
            return _bindingService.DescribeStartupState();
        }

        internal KernelSettingsState LoadSettings()
        {
            return _bindingService.LoadSettings();
        }

        internal void PrepareForHomeDisplay()
        {
            _displayService.PrepareForHomeDisplay();
        }

        internal void PrepareForHomeDisplayFromSheet()
        {
            _displayService.PrepareForHomeDisplayFromSheet();
        }

        internal void CompleteHomeNavigation(bool showExcel)
        {
            _displayService.CompleteHomeNavigation(showExcel);
        }

        internal void EnsureHomeDisplayHidden(string triggerReason)
        {
            _displayService.EnsureHomeDisplayHidden(triggerReason);
        }

        internal bool SaveNameRuleA(string ruleA)
        {
            return _bindingService.SaveNameRuleA(ruleA);
        }

        internal bool SaveNameRuleB(string ruleB)
        {
            return _bindingService.SaveNameRuleB(ruleB);
        }

        internal string SelectAndSaveDefaultRoot()
        {
            return _bindingService.SelectAndSaveDefaultRoot();
        }

        internal void CloseHomeSession()
        {
            _closeService.CloseHomeSession();
        }

        internal void CloseHomeSessionSavingKernel()
        {
            _closeService.CloseHomeSessionSavingKernel();
        }

        private Excel.Workbook ResolveWorkbookForHomeDisplayOrClose(string operationName)
        {
            return _bindingService.ResolveWorkbookForHomeDisplayOrClose(operationName);
        }

        private void ShowKernelWorkbookWindows(bool activateWorkbookWindow)
        {
            _displayService.ShowKernelWorkbookWindows(activateWorkbookWindow);
        }

        private void HideExcelMainWindow()
        {
            _displayService.HideExcelMainWindow();
        }

        private void ShowExcelMainWindow()
        {
            _displayService.ShowExcelMainWindow();
        }

        private void CloseKernelWorkbookWithoutLifecycleCore(Excel.Workbook workbook)
        {
            _closeService.CloseKernelWorkbookWithoutLifecycleCore(workbook);
        }

        private void SaveAndCloseKernelWorkbook(Excel.Workbook workbook)
        {
            _closeService.SaveAndCloseKernelWorkbook(workbook);
        }

        private void QuitApplicationCore()
        {
            _closeService.QuitApplicationCore();
        }

        private static KernelWorkbookFacadeServices CreateFacadeServices(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger)
        {
            var pathCompatibilityService = new PathCompatibilityService();
            var bindingService = new KernelWorkbookBindingService(
                application,
                excelInteropService,
                pathCompatibilityService,
                logger);
            var displayService = new KernelWorkbookDisplayService(
                application,
                excelInteropService,
                excelWindowRecoveryService,
                kernelCaseInteractionState,
                logger,
                bindingService);
            var closeService = new KernelWorkbookCloseService(
                application,
                kernelCaseInteractionState,
                logger,
                bindingService,
                displayService);
            return new KernelWorkbookFacadeServices(bindingService, displayService, closeService);
        }

        private static KernelWorkbookFacadeServices CreateFacadeServices(
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger,
            KernelWorkbookServiceTestHooks testHooks)
        {
            var pathCompatibilityService = new PathCompatibilityService();
            var bindingService = new KernelWorkbookBindingService(
                application: null,
                excelInteropService: null,
                pathCompatibilityService,
                logger,
                testHooks);
            var displayService = new KernelWorkbookDisplayService(
                application: null,
                excelInteropService: null,
                excelWindowRecoveryService: null,
                kernelCaseInteractionState,
                logger,
                bindingService,
                testHooks);
            var closeService = new KernelWorkbookCloseService(
                application: null,
                kernelCaseInteractionState,
                logger,
                bindingService,
                displayService,
                testHooks);
            return new KernelWorkbookFacadeServices(bindingService, displayService, closeService);
        }

        private KernelWorkbookService(KernelWorkbookFacadeServices services)
            : this(services.BindingService, services.DisplayService, services.CloseService)
        {
        }

        private sealed class KernelWorkbookFacadeServices
        {
            internal KernelWorkbookFacadeServices(
                KernelWorkbookBindingService bindingService,
                KernelWorkbookDisplayService displayService,
                KernelWorkbookCloseService closeService)
            {
                BindingService = bindingService;
                DisplayService = displayService;
                CloseService = closeService;
            }

            internal KernelWorkbookBindingService BindingService { get; }

            internal KernelWorkbookDisplayService DisplayService { get; }

            internal KernelWorkbookCloseService CloseService { get; }
        }

        internal sealed class KernelWorkbookServiceTestHooks
        {
            internal Action ApplyHomeDisplayVisibility { get; set; }

            internal Func<string, string> ResolveKernelWorkbookPath { get; set; }

            internal Func<string, Excel.Workbook> FindOpenWorkbook { get; set; }

            internal Func<Excel.Workbook, bool> HasOtherVisibleWorkbook { get; set; }

            internal Func<Excel.Workbook, bool> HasOtherWorkbook { get; set; }

            internal Action<bool> ReleaseHomeDisplay { get; set; }

            internal Action<string> DismissPreparedHomeDisplayState { get; set; }

            internal Action QuitApplication { get; set; }

            internal Func<Excel.Workbook, bool> RequestManagedCloseFromHomeExit { get; set; }

            internal Action<Excel.Workbook> SaveAndCloseKernelWorkbook { get; set; }

            internal Action<Excel.Workbook> CloseKernelWorkbookWithoutLifecycle { get; set; }

            internal Action<Excel.Workbook> ConcealKernelWorkbookWindowsForCaseCreationClose { get; set; }
        }
    }

    internal enum KernelHomeSessionCloseRequestStatus
    {
        Rejected = 0,
        Pending = 1,
        Completed = 2
    }

    internal sealed class KernelSettingsState
    {
        internal string SystemRoot { get; set; } = string.Empty;
        internal string DefaultRoot { get; set; } = string.Empty;
        internal string NameRuleA { get; set; } = "YYYY";
        internal string NameRuleB { get; set; } = "DOC";
    }
}
