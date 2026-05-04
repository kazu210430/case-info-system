using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private const int SwHide = 0;
        private const int SwShow = 5;
        private const int SwRestore = 9;

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly ExcelWindowRecoveryService _excelWindowRecoveryService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly KernelCaseInteractionState _kernelCaseInteractionState;
        private readonly Logger _logger;
        private readonly KernelWorkbookServiceTestHooks _testHooks;
        private readonly KernelOpenWorkbookLocator _kernelOpenWorkbookLocator;
        private readonly KernelWorkbookStateService _kernelWorkbookStateService;
        private readonly KernelWorkbookSettingsService _kernelWorkbookSettingsService;
        private readonly KernelHomeSessionCloseCoordinator _homeSessionCloseCoordinator;
        private KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private bool _isHomeDisplayPrepared;
        private KernelHomeBinding _homeBinding;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);


        /// <summary>
        /// メソッド: Kernel Workbook の close 制御に利用する lifecycle service を設定する。
        /// 引数: kernelWorkbookLifecycleService - lifecycle service。
        /// 戻り値: なし。
        /// 副作用: HOME 終了時の close 経路で利用する参照を保持する。
        /// </summary>
        internal void SetLifecycleService(KernelWorkbookLifecycleService kernelWorkbookLifecycleService)
        {
            _kernelWorkbookLifecycleService = kernelWorkbookLifecycleService ?? throw new ArgumentNullException(nameof(kernelWorkbookLifecycleService));
        }
        internal KernelWorkbookService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _excelWindowRecoveryService = excelWindowRecoveryService ?? throw new ArgumentNullException(nameof(excelWindowRecoveryService));
            _pathCompatibilityService = new PathCompatibilityService();
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _testHooks = null;
            _kernelOpenWorkbookLocator = new KernelOpenWorkbookLocator(_application, _excelInteropService, _pathCompatibilityService, _logger);
            _kernelWorkbookStateService = new KernelWorkbookStateService(_application, _excelInteropService, _logger, _kernelOpenWorkbookLocator);
            _kernelWorkbookSettingsService = new KernelWorkbookSettingsService();
            _homeSessionCloseCoordinator = new KernelHomeSessionCloseCoordinator(this);
        }

        internal KernelWorkbookService(KernelCaseInteractionState kernelCaseInteractionState, Logger logger, KernelWorkbookServiceTestHooks testHooks)
        {
            _application = null;
            _excelInteropService = null;
            _excelWindowRecoveryService = null;
            _pathCompatibilityService = new PathCompatibilityService();
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _testHooks = testHooks;
            _kernelOpenWorkbookLocator = new KernelOpenWorkbookLocator(
                _application,
                _excelInteropService,
                _pathCompatibilityService,
                _logger,
                getOpenKernelWorkbookOverride: _testHooks == null ? null : _testHooks.GetOpenKernelWorkbook,
                resolveKernelWorkbookPathOverride: _testHooks == null ? null : _testHooks.ResolveKernelWorkbookPath,
                findOpenWorkbookOverride: _testHooks == null ? null : _testHooks.FindOpenWorkbook);
            _kernelWorkbookStateService = new KernelWorkbookStateService(
                _application,
                _excelInteropService,
                _logger,
                _kernelOpenWorkbookLocator,
                hasOtherVisibleWorkbookOverride: _testHooks == null ? null : _testHooks.HasOtherVisibleWorkbook,
                hasOtherWorkbookOverride: _testHooks == null ? null : _testHooks.HasOtherWorkbook);
            _kernelWorkbookSettingsService = new KernelWorkbookSettingsService();
            _homeSessionCloseCoordinator = new KernelHomeSessionCloseCoordinator(this);
        }

        internal Excel.Workbook GetOpenKernelWorkbook()
        {
            return _kernelOpenWorkbookLocator.GetOpenKernelWorkbook();
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return _kernelWorkbookStateService.IsKernelWorkbook(workbook);
        }

        internal Excel.Workbook ResolveKernelWorkbook(Domain.WorkbookContext context)
        {
            return _kernelOpenWorkbookLocator.ResolveKernelWorkbook(context);
        }

        internal Excel.Workbook ResolveKernelWorkbook(string systemRoot)
        {
            return _kernelOpenWorkbookLocator.ResolveKernelWorkbook(systemRoot);
        }

        internal bool BindHomeWorkbook(Domain.WorkbookContext context)
        {
            KernelHomeBinding binding = CreateHomeBinding(context);
            if (binding == null)
            {
                ClearHomeWorkbookBinding("BindHomeWorkbook.Invalid");
                _logger.Warn(
                    "Kernel HOME binding was not created from context. workbook="
                    + GetWorkbookFullNameCore(context == null ? null : context.Workbook)
                    + ", contextSystemRoot="
                    + (context == null ? string.Empty : context.SystemRoot ?? string.Empty));
                return false;
            }

            _homeBinding = binding;
            _logger.Info(
                "Kernel HOME binding created. workbook="
                + GetWorkbookFullNameCore(binding.Workbook)
                + ", systemRoot="
                + binding.SystemRoot);
            return true;
        }

        internal bool BindHomeWorkbook(Excel.Workbook workbook)
        {
            KernelHomeBinding binding = CreateHomeBinding(workbook);
            if (binding == null)
            {
                ClearHomeWorkbookBinding("BindHomeWorkbook.WorkbookInvalid");
                _logger.Warn(
                    "Kernel HOME binding was not created from workbook. workbook="
                    + GetWorkbookFullNameCore(workbook));
                return false;
            }

            _homeBinding = binding;
            _logger.Info(
                "Kernel HOME binding created from workbook. workbook="
                + GetWorkbookFullNameCore(binding.Workbook)
                + ", systemRoot="
                + binding.SystemRoot);
            return true;
        }

        internal void ClearHomeWorkbookBinding(string reason)
        {
            if (_homeBinding == null)
            {
                return;
            }

            _logger.Info(
                "Kernel HOME binding cleared. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + GetWorkbookFullNameCore(_homeBinding.Workbook)
                + ", systemRoot="
                + (_homeBinding.SystemRoot ?? string.Empty));
            _homeBinding = null;
        }

        internal bool HasValidHomeWorkbookBinding()
        {
            Excel.Workbook workbook;
            return ResolveHomeBindingStatus("HasValidHomeWorkbookBinding", out workbook) == KernelHomeBindingStatus.Valid;
        }

        internal bool TryGetValidHomeWorkbookBinding(out Excel.Workbook workbook, out string systemRoot)
        {
            workbook = null;
            systemRoot = string.Empty;

            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus("TryGetValidHomeWorkbookBinding", out workbook);
            if (bindingStatus != KernelHomeBindingStatus.Valid)
            {
                workbook = null;
                return false;
            }

            systemRoot = _homeBinding == null ? string.Empty : _homeBinding.SystemRoot ?? string.Empty;
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                workbook = null;
                return false;
            }

            return true;
        }

        internal bool TryShowSheetByCodeName(Domain.WorkbookContext context, string sheetCodeName, string reason)
        {
            Excel.Workbook kernelWorkbook = ResolveKernelWorkbook(context);
            if (kernelWorkbook == null)
            {
                _logger.Warn("TryShowSheetByCodeName skipped because kernel workbook was not available. reason=" + (reason ?? string.Empty));
                return false;
            }

            bool activated = _excelInteropService.ActivateWorkbook(kernelWorkbook);
            bool sheetActivated = activated && _excelInteropService.ActivateWorksheetByCodeName(kernelWorkbook, sheetCodeName);
            _logger.Info(
                "TryShowSheetByCodeName result=" + sheetActivated.ToString()
                + ", reason=" + (reason ?? string.Empty)
                + ", sheetCodeName=" + (sheetCodeName ?? string.Empty));
            return sheetActivated;
        }

        internal bool ShouldShowHomeOnStartup(Excel.Workbook startupWorkbook = null)
        {
            return _kernelWorkbookStateService.ShouldShowHomeOnStartup(startupWorkbook);
        }

        internal string DescribeStartupState()
        {
            return _kernelWorkbookStateService.DescribeStartupState();
        }

        internal KernelSettingsState LoadSettings()
        {
            Excel.Workbook workbook;
            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus("LoadSettings", out workbook);
            if (bindingStatus == KernelHomeBindingStatus.Invalid)
            {
                _logger.Warn("Kernel settings load failed closed because HOME binding was invalid.");
                return new KernelSettingsState();
            }

            if (bindingStatus == KernelHomeBindingStatus.None)
            {
                workbook = GetOrOpenKernelWorkbook();
            }

            if (workbook == null)
            {
                _logger.Warn("Kernel settings load skipped because kernel workbook was not available.");
                return new KernelSettingsState();
            }

            string nameRuleA = _kernelWorkbookSettingsService.LoadNameRuleA(workbook);
            string nameRuleB = _kernelWorkbookSettingsService.LoadNameRuleB(workbook);
            string defaultRoot = _kernelWorkbookSettingsService.LoadDefaultRoot(workbook);
            string systemRoot = KernelWorkbookResolver.TryGetKernelWorkbookDirectory(workbook, _excelInteropService, _pathCompatibilityService, _logger, IsKernelWorkbook);

            _logger.Info(
                "Kernel settings loaded. workbook="
                + _excelInteropService.GetWorkbookFullName(workbook)
                + ", systemRoot="
                + (systemRoot ?? string.Empty)
                + ", defaultRoot="
                + (defaultRoot ?? string.Empty)
                + ", nameRuleA="
                + (nameRuleA ?? string.Empty)
                + ", nameRuleB="
                + (nameRuleB ?? string.Empty));

            return new KernelSettingsState
            {
                SystemRoot = systemRoot,
                DefaultRoot = defaultRoot,
                NameRuleA = nameRuleA,
                NameRuleB = nameRuleB
            };
        }

        internal void PrepareForHomeDisplay()
        {
            if (_isHomeDisplayPrepared)
            {
                return;
            }

            ApplyHomeDisplayVisibilityCore("PrepareForHomeDisplay");
            _isHomeDisplayPrepared = true;
        }

        internal void PrepareForHomeDisplayFromSheet()
        {
            ApplyHomeDisplayVisibilityCore("PrepareForHomeDisplayFromSheet");
            _isHomeDisplayPrepared = true;
        }

        internal void CompleteHomeNavigation(bool showExcel)
        {
            ReleaseHomeDisplay(showExcel);
        }

        internal void EnsureHomeDisplayHidden(string triggerReason)
        {
            string caller = ResolveExternalCaller();
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=ensure-home-display-hidden-enter trigger="
                + (triggerReason ?? string.Empty)
                + ", caller="
                + caller
                + ", activeState="
                + FormatActiveExcelState()
                + ", isHomeDisplayPrepared="
                + _isHomeDisplayPrepared.ToString()
                + ", tracePresent="
                + (!string.IsNullOrWhiteSpace(KernelFlickerTraceContext.CurrentTraceId)).ToString());
            if (!_isHomeDisplayPrepared)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=ensure-home-display-hidden-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=skipped-not-prepared");
                return;
            }

            ApplyHomeDisplayVisibilityCore("EnsureHomeDisplayHidden|" + (triggerReason ?? string.Empty));
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=ensure-home-display-hidden-end trigger="
                + (triggerReason ?? string.Empty)
                + ", result=applied, activeState="
                + FormatActiveExcelState());
        }

        internal bool SaveNameRuleA(string ruleA)
        {
            Excel.Workbook workbook = ResolveWorkbookForBoundHomeMutation("SaveNameRuleA");
            if (workbook == null)
            {
                return false;
            }

            _kernelWorkbookSettingsService.SaveNameRuleA(workbook, ruleA);
            workbook.Save();
            return true;
        }

        internal bool SaveNameRuleB(string ruleB)
        {
            Excel.Workbook workbook = ResolveWorkbookForBoundHomeMutation("SaveNameRuleB");
            if (workbook == null)
            {
                return false;
            }

            _kernelWorkbookSettingsService.SaveNameRuleB(workbook, ruleB);
            workbook.Save();
            return true;
        }

        internal string SelectAndSaveDefaultRoot()
        {
            Excel.Workbook workbook = ResolveWorkbookForBoundHomeMutation("SelectAndSaveDefaultRoot");
            if (workbook == null)
            {
                return null;
            }

            string selectedPath = _kernelWorkbookSettingsService.SelectDefaultRoot(workbook);
            if (string.IsNullOrWhiteSpace(selectedPath))
            {
                return null;
            }

            workbook.Save();
            return selectedPath;
        }

        internal void CloseHomeSession()
        {
            CloseHomeSession(saveKernelWorkbook: false, entryPoint: "CloseHomeSession");
        }

        internal void CloseHomeSessionSavingKernel()
        {
            CloseHomeSession(saveKernelWorkbook: true, entryPoint: "CloseHomeSessionSavingKernel");
        }

        private void CloseHomeSession(bool saveKernelWorkbook, string entryPoint)
        {
            string caller = ResolveExternalCaller();
            _homeSessionCloseCoordinator.Execute(saveKernelWorkbook, entryPoint, caller);
        }

        private Excel.Workbook GetOpenKernelWorkbookCore()
        {
            return _testHooks != null && _testHooks.GetOpenKernelWorkbook != null
                ? _testHooks.GetOpenKernelWorkbook()
                : GetOpenKernelWorkbook();
        }

        private KernelHomeBinding CreateHomeBinding(Domain.WorkbookContext context)
        {
            if (context == null)
            {
                return null;
            }

            string normalizedContextSystemRoot = _pathCompatibilityService.NormalizePath(context.SystemRoot);
            if (string.IsNullOrWhiteSpace(normalizedContextSystemRoot))
            {
                return null;
            }

            return CreateHomeBinding(context.Workbook, normalizedContextSystemRoot);
        }

        private KernelHomeBinding CreateHomeBinding(Excel.Workbook workbook)
        {
            return CreateHomeBinding(workbook, systemRoot: null);
        }

        private KernelHomeBinding CreateHomeBinding(Excel.Workbook workbook, string systemRoot)
        {
            if (!IsKernelWorkbook(workbook))
            {
                return null;
            }

            string workbookSystemRoot = GetWorkbookSystemRootForHomeBinding(workbook);
            if (string.IsNullOrWhiteSpace(workbookSystemRoot))
            {
                return null;
            }

            string normalizedSystemRoot = _pathCompatibilityService.NormalizePath(systemRoot);
            if (!string.IsNullOrWhiteSpace(normalizedSystemRoot)
                && !string.Equals(workbookSystemRoot, normalizedSystemRoot, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return new KernelHomeBinding(workbook, string.IsNullOrWhiteSpace(normalizedSystemRoot) ? workbookSystemRoot : normalizedSystemRoot);
        }

        private KernelHomeBindingStatus ResolveHomeBindingStatus(string operationName, out Excel.Workbook workbook)
        {
            workbook = null;
            if (_homeBinding == null)
            {
                return KernelHomeBindingStatus.None;
            }

            Excel.Workbook boundWorkbook = _homeBinding.Workbook;
            if (!IsKernelWorkbook(boundWorkbook))
            {
                _logger.Warn(
                    "Kernel HOME binding became invalid because workbook was not kernel. operation="
                    + (operationName ?? string.Empty)
                    + ", workbook="
                    + GetWorkbookFullNameCore(boundWorkbook));
                return KernelHomeBindingStatus.Invalid;
            }

            string currentSystemRoot = GetWorkbookSystemRootForHomeBinding(boundWorkbook);
            if (string.IsNullOrWhiteSpace(currentSystemRoot)
                || !string.Equals(currentSystemRoot, _homeBinding.SystemRoot, StringComparison.OrdinalIgnoreCase))
            {
                _logger.Warn(
                    "Kernel HOME binding became invalid because system root mismatched. operation="
                    + (operationName ?? string.Empty)
                    + ", workbook="
                    + GetWorkbookFullNameCore(boundWorkbook)
                    + ", boundSystemRoot="
                    + (_homeBinding.SystemRoot ?? string.Empty)
                    + ", currentSystemRoot="
                    + (currentSystemRoot ?? string.Empty));
                return KernelHomeBindingStatus.Invalid;
            }

            workbook = boundWorkbook;
            return KernelHomeBindingStatus.Valid;
        }

        private Excel.Workbook ResolveWorkbookForBoundHomeMutation(string operationName)
        {
            Excel.Workbook workbook;
            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus(operationName, out workbook);
            if (bindingStatus != KernelHomeBindingStatus.Valid)
            {
                _logger.Warn(
                    "Kernel HOME mutation failed closed because valid binding was not available. operation="
                    + (operationName ?? string.Empty)
                    + ", bindingStatus="
                    + bindingStatus.ToString());
                return null;
            }

            return workbook;
        }

        private Excel.Workbook ResolveWorkbookForHomeDisplayOrClose(string operationName)
        {
            Excel.Workbook workbook;
            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus(operationName, out workbook);
            if (bindingStatus == KernelHomeBindingStatus.Valid)
            {
                return workbook;
            }

            if (bindingStatus == KernelHomeBindingStatus.Invalid)
            {
                return null;
            }

            return GetOpenKernelWorkbookCore();
        }

        private string GetWorkbookSystemRootForHomeBinding(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            if (_excelInteropService != null)
            {
                return _pathCompatibilityService.NormalizePath(
                    KernelWorkbookResolver.GetSystemRootFromWorkbook(
                        workbook,
                        _excelInteropService,
                        _pathCompatibilityService,
                        _logger,
                        IsKernelWorkbook));
            }

            try
            {
                if (workbook.CustomDocumentProperties is IDictionary<string, string> properties
                    && properties.TryGetValue("SYSTEM_ROOT", out string systemRoot))
                {
                    return _pathCompatibilityService.NormalizePath(systemRoot);
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private bool HasOtherVisibleWorkbookCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.HasOtherVisibleWorkbook != null
                ? _testHooks.HasOtherVisibleWorkbook(workbook)
                : HasOtherVisibleWorkbook(workbook);
        }

        private string GetWorkbookFullNameCore(Excel.Workbook workbook)
        {
            if (_excelInteropService != null)
            {
                return _excelInteropService.GetWorkbookFullName(workbook);
            }

            return workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
        }

        private bool HasOtherWorkbookCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.HasOtherWorkbook != null
                ? _testHooks.HasOtherWorkbook(workbook)
                : HasOtherWorkbook(workbook);
        }

        private void ReleaseHomeDisplayCore(bool showExcel)
        {
            if (_testHooks != null && _testHooks.ReleaseHomeDisplay != null)
            {
                _testHooks.ReleaseHomeDisplay(showExcel);
                return;
            }

            ReleaseHomeDisplay(showExcel);
        }

        private void DismissPreparedHomeDisplayStateCore(string reason)
        {
            if (_testHooks != null && _testHooks.DismissPreparedHomeDisplayState != null)
            {
                _testHooks.DismissPreparedHomeDisplayState(reason);
                return;
            }

            DismissPreparedHomeDisplayState(reason);
        }

        private void QuitApplicationCore()
        {
            if (_testHooks != null && _testHooks.QuitApplication != null)
            {
                _testHooks.QuitApplication();
                return;
            }

            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                _application.DisplayAlerts = false;
                _application.Quit();
            }
            finally
            {
                _application.DisplayAlerts = previousDisplayAlerts;
            }
        }

        private bool RequestManagedCloseFromHomeExitCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.RequestManagedCloseFromHomeExit != null
                ? _testHooks.RequestManagedCloseFromHomeExit(workbook)
                : _kernelWorkbookLifecycleService.RequestManagedCloseFromHomeExit(workbook);
        }

        private void SaveAndCloseKernelWorkbookCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.SaveAndCloseKernelWorkbook != null)
            {
                _testHooks.SaveAndCloseKernelWorkbook(workbook);
                return;
            }

            SaveAndCloseKernelWorkbook(workbook);
        }

        private void CloseKernelWorkbookWithoutLifecycleCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.CloseKernelWorkbookWithoutLifecycle != null)
            {
                _testHooks.CloseKernelWorkbookWithoutLifecycle(workbook);
                return;
            }

            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                _application.DisplayAlerts = false;
                workbook.Close(SaveChanges: false);
            }
            finally
            {
                _application.DisplayAlerts = previousDisplayAlerts;
            }
        }

        private void ConcealKernelWorkbookWindowsForCaseCreationCloseCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.ConcealKernelWorkbookWindowsForCaseCreationClose != null)
            {
                _testHooks.ConcealKernelWorkbookWindowsForCaseCreationClose(workbook);
                return;
            }

            ConcealKernelWorkbookWindowsForCaseCreationClose(workbook);
        }

        private void ApplyHomeDisplayVisibilityCore(string triggerReason)
        {
            if (_testHooks != null && _testHooks.ApplyHomeDisplayVisibility != null)
            {
                _testHooks.ApplyHomeDisplayVisibility();
                return;
            }

            ApplyHomeDisplayVisibility(triggerReason);
        }

        internal bool ShowSheetByCodeName(string codeName)
        {
            Excel.Workbook displayedWorkbook;
            return ShowSheetByCodeName(codeName, out displayedWorkbook);
        }

        internal bool ShowSheetByCodeName(string codeName, out Excel.Workbook displayedWorkbook)
        {
            displayedWorkbook = null;
            Excel.Workbook workbook = GetOrOpenKernelWorkbook();
            if (workbook == null)
            {
                return false;
            }

            bool previousEnableEvents = _application.EnableEvents;
            try
            {
                _application.EnableEvents = false;
                PrepareWorkbookForSheetNavigation(workbook, codeName);
                Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName(workbook, codeName);
                if (worksheet == null)
                {
                    return false;
                }

                workbook.Activate();
                worksheet.Activate();
                displayedWorkbook = workbook;
                return true;
            }
            finally
            {
                _application.EnableEvents = previousEnableEvents;
            }
        }

        private void PrepareWorkbookForSheetNavigation(Excel.Workbook workbook, string codeName)
        {
            if (_isHomeDisplayPrepared)
            {
                DismissPreparedHomeDisplayState("ShowSheetByCodeName:" + (codeName ?? string.Empty));
            }

            EnsureWorkbookVisible(workbook);
        }

        private Excel.Workbook GetOrOpenKernelWorkbook()
        {
            Excel.Workbook workbook = GetOpenKernelWorkbook();
            if (workbook != null)
            {
                if (_isHomeDisplayPrepared)
                {
                    HideKernelWorkbookWindows(workbook);
                }

                return workbook;
            }

            string workbookPath = KernelWorkbookResolver.ResolveKernelWorkbookPathFromAvailableWorkbooks(
                _application,
                _excelInteropService,
                _pathCompatibilityService,
                _logger,
                IsKernelWorkbook);
            if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
            {
                return null;
            }

            bool previousEnableEvents = _application.EnableEvents;
            try
            {
                _application.EnableEvents = false;
                workbook = _application.Workbooks.Open(workbookPath, ReadOnly: false);
                if (workbook.Windows.Count > 0)
                {
                    SetKernelWindowVisibleFalse(
                        workbook,
                        workbook.Windows[1],
                        1,
                        "GetOrOpenKernelWorkbook.OpenedWorkbook");
                }

                if (_isHomeDisplayPrepared)
                {
                    HideKernelWorkbookWindows(workbook);
                }

                return workbook;
            }
            finally
            {
                _application.EnableEvents = previousEnableEvents;
            }
        }

        private void SaveAndCloseKernelWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
                bool requiresSave = RequiresSave(workbook);
                _logger.Info(
                    "SaveAndCloseKernelWorkbook started. workbook="
                    + workbookFullName
                    + ", requiresSave="
                    + requiresSave.ToString());
                _application.DisplayAlerts = false;

                if (requiresSave)
                {
                    workbook.Save();
                    _logger.Info("SaveAndCloseKernelWorkbook saved workbook=" + workbookFullName);
                }
                else
                {
                    _logger.Info("SaveAndCloseKernelWorkbook skipped save because workbook was already saved. workbook=" + workbookFullName);
                }

                workbook.Close(SaveChanges: false);
                _logger.Info("SaveAndCloseKernelWorkbook closed workbook=" + workbookFullName);
            }
            finally
            {
                _application.DisplayAlerts = previousDisplayAlerts;
            }
        }

        /// <summary>
        /// メソッド: CASE 作成フローで Kernel を閉じる直前に、Kernel workbook window を画面から外す。
        /// 引数: workbook - 閉じる対象の Kernel workbook。
        /// 戻り値: なし。
        /// 副作用: 対象 window を不可視化する。
        /// </summary>
        private void ConcealKernelWorkbookWindowsForCaseCreationClose(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        if (window == null)
                        {
                            continue;
                        }

                        SetKernelWindowVisibleFalse(
                            workbook,
                            window,
                            index,
                            "ConcealKernelWorkbookWindowsForCaseCreationClose");
                    }
                    catch (Exception ex)
                    {
                        // 例外処理: close 継続を優先するため、不可視化失敗はログ化のみで握りつぶす。
                        _logger.Error("ConcealKernelWorkbookWindowsForCaseCreationClose window conceal failed. index=" + index.ToString(), ex);
                    }
                }

                _logger.Info("ConcealKernelWorkbookWindowsForCaseCreationClose completed. workbook=" + _excelInteropService.GetWorkbookFullName(workbook));
            }
            catch (Exception ex)
            {
                // 例外処理: close 継続を優先するため、不可視化全体失敗はログ化のみで握りつぶす。
                _logger.Error("ConcealKernelWorkbookWindowsForCaseCreationClose failed.", ex);
            }
        }

        private static bool RequiresSave(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                return !workbook.Saved;
            }
            catch
            {
                return true;
            }
        }

        private bool HasOtherVisibleWorkbook(Excel.Workbook workbookToIgnore)
        {
            return _kernelWorkbookStateService.HasOtherVisibleWorkbook(workbookToIgnore);
        }

        private bool HasOtherWorkbook(Excel.Workbook workbookToIgnore)
        {
            return _kernelWorkbookStateService.HasOtherWorkbook(workbookToIgnore);
        }

        private bool HasVisibleNonKernelWorkbook()
        {
            return _kernelWorkbookStateService.HasVisibleNonKernelWorkbook();
        }

        private void EnsureWorkbookVisible(Excel.Workbook workbook)
        {
            ReleaseHomeDisplay(true);
            bool shouldAvoidGlobalRestore = ShouldAvoidGlobalExcelWindowRestore();
            if (!shouldAvoidGlobalRestore)
            {
                _excelWindowRecoveryService.EnsureApplicationVisible("KernelWorkbookService.EnsureWorkbookVisible", _excelInteropService.GetWorkbookFullName(workbook));
            }

            _excelWindowRecoveryService.TryRecoverWorkbookWindow(
                workbook,
                "KernelWorkbookService.EnsureWorkbookVisible",
                bringToFront: !shouldAvoidGlobalRestore);
        }

        private void ReleaseHomeDisplay(bool showExcel)
        {
            if (!_isHomeDisplayPrepared)
            {
                return;
            }

            if (showExcel)
            {
                bool shouldAvoidGlobalExcelWindowRestore = ShouldAvoidGlobalExcelWindowRestore();
                bool shouldPromoteKernelWindow = !shouldAvoidGlobalExcelWindowRestore
                    && ShouldPromoteKernelWorkbookOnHomeRelease();
                KernelWorkbookHomeReleaseAction homeReleaseAction = KernelWorkbookHomeReleaseFallbackPolicy.DecideHomeReleaseAction(
                    shouldAvoidGlobalExcelWindowRestore: shouldAvoidGlobalExcelWindowRestore,
                    shouldPromoteKernelWorkbook: shouldPromoteKernelWindow);

                if (homeReleaseAction == KernelWorkbookHomeReleaseAction.SkipRestore)
                {
                    _logger.Info("ReleaseHomeDisplay skipped global Excel window restore to preserve other workbook layouts.");
                    _isHomeDisplayPrepared = false;
                    return;
                }

                if (homeReleaseAction == KernelWorkbookHomeReleaseAction.PromoteAndRestore)
                {
                    ShowExcelMainWindow();
                }

                ShowKernelWorkbookWindows(homeReleaseAction == KernelWorkbookHomeReleaseAction.PromoteAndRestore);
            }

            _isHomeDisplayPrepared = false;
        }

        /// <summary>
        /// メソッド: HOME 表示準備状態だけを解除し、Excel/Kernel の表示復帰は行わない。
        /// 引数: reason - ログ出力用の理由。
        /// 戻り値: なし。
        /// 副作用: HOME 表示準備フラグを解除する。
        /// </summary>
        private void DismissPreparedHomeDisplayState(string reason)
        {
            if (!_isHomeDisplayPrepared)
            {
                return;
            }

            _isHomeDisplayPrepared = false;
            _logger.Info("DismissPreparedHomeDisplayState executed. reason=" + (reason ?? string.Empty));
        }

        private void ApplyHomeDisplayVisibility(string triggerReason)
        {
            Excel.Workbook kernelWorkbook = ResolveWorkbookForHomeDisplayOrClose("ApplyHomeDisplayVisibility");
            bool hasVisibleNonKernelWorkbook = HasVisibleNonKernelWorkbook();
            bool preserveOtherWorkbookWindowLayout = ShouldPreserveOtherWorkbookWindowLayout();
            string visibleNonKernelWindows = DescribeVisibleNonKernelWorkbookWindows();
            string kernelWindowTargets = DescribeWorkbookWindows(kernelWorkbook);
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=apply-home-display-enter trigger="
                + (triggerReason ?? string.Empty)
                + ", activeState="
                + FormatActiveExcelState()
                + ", hasVisibleNonKernelWorkbook="
                + hasVisibleNonKernelWorkbook.ToString()
                + ", preserveOtherWorkbookWindowLayout="
                + preserveOtherWorkbookWindowLayout.ToString()
                + ", visibleNonKernelWindows="
                + visibleNonKernelWindows
                + ", kernelWindowTargets="
                + kernelWindowTargets);
            if (hasVisibleNonKernelWorkbook)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=apply-home-display-decision trigger="
                    + (triggerReason ?? string.Empty)
                    + ", decision=minimize-kernel-windows, reason=visible-non-kernel-workbook-detected, preserveOtherWorkbookWindowLayout="
                    + preserveOtherWorkbookWindowLayout.ToString()
                    + ", visibleNonKernelWindows="
                    + visibleNonKernelWindows
                    + ", kernelWindowTargets="
                    + kernelWindowTargets);
                if (!preserveOtherWorkbookWindowLayout)
                {
                    EnsureExcelApplicationVisible();
                }

                HideKernelWorkbookWindows("ApplyHomeDisplayVisibility:" + (triggerReason ?? string.Empty), kernelWorkbook);
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=apply-home-display-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=minimized-kernel-windows, visibleNonKernelWindows="
                    + visibleNonKernelWindows
                    + ", kernelWindowTargets="
                    + kernelWindowTargets);
                _logger.Info(
                    "ApplyHomeDisplayVisibility minimized kernel windows because a non-kernel workbook is visible. preserveOtherWindowLayout="
                    + preserveOtherWorkbookWindowLayout.ToString());
                return;
            }

            Excel.Workbook activeWorkbook = null;
            bool shouldSkipHideExcelMainWindow = false;
            try
            {
                activeWorkbook = _application == null ? null : _application.ActiveWorkbook;
                shouldSkipHideExcelMainWindow = activeWorkbook != null
                    && IsKernelWorkbook(activeWorkbook)
                    && CountVisibleWorkbooksSafe() >= 1;
            }
            catch
            {
                shouldSkipHideExcelMainWindow = false;
            }

            if (shouldSkipHideExcelMainWindow)
            {
                Excel.Workbook workbookToConceal = kernelWorkbook ?? activeWorkbook;
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=apply-home-display-decision trigger="
                    + (triggerReason ?? string.Empty)
                    + ", decision=conceal-kernel-windows-and-hide-excel-main-window, reason=active-kernel-workbook-still-visible, visibleWorkbookCount="
                    + CountVisibleWorkbooksSafe().ToString()
                    + ", activeWorkbook="
                    + FormatWorkbookDescriptor(activeWorkbook)
                    + ", concealTarget="
                    + FormatWorkbookDescriptor(workbookToConceal)
                    + ", kernelWindowTargets="
                    + kernelWindowTargets);
                ConcealKernelWorkbookWindowsForHomeDisplay(workbookToConceal, "ApplyHomeDisplayVisibility:" + (triggerReason ?? string.Empty));
                _logger.Info(
                    "ApplyHomeDisplayVisibility concealed kernel windows before hiding Excel main window because the active kernel workbook remained visible.");
            }

            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=apply-home-display-decision trigger="
                + (triggerReason ?? string.Empty)
                + ", decision=hide-excel-main-window, reason=no-visible-non-kernel-workbook, kernelWindowTargets="
                + kernelWindowTargets);
            HideExcelMainWindow();
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=apply-home-display-end trigger="
                + (triggerReason ?? string.Empty)
                + ", result=excel-main-window-hidden, activeState="
                + FormatActiveExcelState());
        }

        private bool ShouldPreserveOtherWorkbookWindowLayout()
        {
            // 処理ブロック: Kernel→CASE 作成フロー中以外は他 workbook のレイアウト維持を優先する.
            return !_kernelCaseInteractionState.IsKernelCaseCreationFlowActive;
        }

        private bool ShouldAvoidGlobalExcelWindowRestore()
        {
            return KernelWorkbookWindowRestorePolicy.ShouldAvoidGlobalExcelWindowRestore(
                isKernelCaseCreationFlowActive: _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                hasVisibleNonKernelWorkbook: HasVisibleNonKernelWorkbook());
        }

        private void HideExcelMainWindow()
        {
            try
            {
                LogHideExcelMainWindowState("before");
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                ShowWindow(hwnd, SwHide);
                _application.Visible = false;
                LogHideExcelMainWindowState("after");
            }
            catch
            {
            }
        }

        private void LogHideExcelMainWindowState(string stage)
        {
            _logger.Info(
                "HideExcelMainWindow state. stage="
                + (stage ?? string.Empty)
                + ", applicationVisible="
                + SafeApplicationVisible()
                + ", applicationHwnd="
                + SafeApplicationHwnd()
                + ", activeWorkbook="
                + SafeActiveWorkbookDescriptor()
                + ", activeWindow="
                + SafeActiveWindowDescriptor()
                + ", visibleWorkbookCount="
                + CountVisibleWorkbooksSafe().ToString());
        }

        private string SafeApplicationVisible()
        {
            try
            {
                return _application == null ? string.Empty : _application.Visible.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeApplicationHwnd()
        {
            try
            {
                return _application == null ? string.Empty : Convert.ToString(_application.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeActiveWorkbookDescriptor()
        {
            try
            {
                Excel.Workbook workbook = _application == null ? null : _application.ActiveWorkbook;
                return workbook == null ? string.Empty : GetWorkbookFullNameCore(workbook);
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeActiveWindowDescriptor()
        {
            try
            {
                Excel.Window window = _application == null ? null : _application.ActiveWindow;
                return window == null ? string.Empty : FormatWindowDescriptor(window);
            }
            catch
            {
                return string.Empty;
            }
        }

        private int CountVisibleWorkbooksSafe()
        {
            try
            {
                int count = 0;
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (workbook == null || workbook.Windows == null)
                    {
                        continue;
                    }

                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (window != null && window.Visible)
                        {
                            count++;
                            break;
                        }
                    }
                }

                return count;
            }
            catch
            {
                return -1;
            }
        }

        private void HideKernelWorkbookWindows(Excel.Workbook workbook)
        {
            HideKernelWorkbookWindows("HideKernelWorkbookWindows.DirectWorkbook", workbook);
        }

        private void HideKernelWorkbookWindows(string triggerReason, Excel.Workbook workbook)
        {
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=hide-kernel-windows-enter trigger="
                + (triggerReason ?? string.Empty)
                + ", targetWorkbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState()
                + ", targets="
                + DescribeWorkbookWindows(workbook));
            if (workbook == null)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=hide-kernel-windows-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=skipped-null-workbook");
                return;
            }

            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                int minimizedCount = 0;
                int failedCount = 0;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        string beforeState = FormatWindowDescriptor(window);
                        LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=minimize-window-start trigger="
                            + (triggerReason ?? string.Empty)
                            + ", index="
                            + index.ToString()
                            + ", workbook="
                            + FormatWorkbookDescriptor(workbook)
                            + ", window="
                            + beforeState);
                        if (window != null)
                        {
                            bool isVisible = false;
                            try
                            {
                                isVisible = window.Visible;
                            }
                            catch
                            {
                            }

                            if (!isVisible)
                            {
                                LogKernelFlickerTrace(
                                    "source=KernelWorkbookService action=minimize-window-end trigger="
                                    + (triggerReason ?? string.Empty)
                                    + ", index="
                                    + index.ToString()
                                    + ", result=skipped-already-invisible, workbook="
                                    + FormatWorkbookDescriptor(workbook)
                                    + ", window="
                                    + beforeState);
                                continue;
                            }

                            window.WindowState = Excel.XlWindowState.xlMinimized;
                            minimizedCount++;
                            LogKernelFlickerTrace(
                                "source=KernelWorkbookService action=minimize-window-end trigger="
                                + (triggerReason ?? string.Empty)
                                + ", index="
                                + index.ToString()
                                + ", result=success, workbook="
                                + FormatWorkbookDescriptor(workbook)
                                + ", windowBefore="
                                + beforeState
                                + ", windowAfter="
                                + FormatWindowDescriptor(window));
                        }
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=minimize-window-end trigger="
                            + (triggerReason ?? string.Empty)
                            + ", index="
                            + index.ToString()
                            + ", result=failed, workbook="
                            + FormatWorkbookDescriptor(workbook)
                            + ", window="
                            + FormatWindowDescriptor(window)
                            + ", exceptionType="
                            + ex.GetType().Name
                            + ", exceptionMessage="
                            + (ex.Message ?? string.Empty));
                        _logger.Error("HideKernelWorkbookWindows window minimize failed. index=" + index.ToString(), ex);
                    }
                }

                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=hide-kernel-windows-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=completed, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", totalTargets="
                    + windowCount.ToString()
                    + ", minimizedCount="
                    + minimizedCount.ToString()
                    + ", failedCount="
                    + failedCount.ToString());
            }
            catch (Exception ex)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=hide-kernel-windows-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=failed, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", exceptionType="
                    + ex.GetType().Name
                    + ", exceptionMessage="
                    + (ex.Message ?? string.Empty));
                _logger.Error("HideKernelWorkbookWindows failed.", ex);
            }
        }

        private void ShowKernelWorkbookWindows(bool activateWorkbookWindow)
        {
            Excel.Workbook workbook = GetOpenKernelWorkbook();
            if (workbook == null)
            {
                return;
            }

            try
            {
                _excelWindowRecoveryService.EnsureApplicationVisible("KernelWorkbookService.ShowKernelWorkbookWindows", _excelInteropService.GetWorkbookFullName(workbook));
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null)
                    {
                        window.Visible = true;
                        window.WindowState = Excel.XlWindowState.xlNormal;
                    }
                }

                if (activateWorkbookWindow)
                {
                    _excelWindowRecoveryService.TryRecoverWorkbookWindow(
                        workbook,
                        "KernelWorkbookService.ShowKernelWorkbookWindows",
                        bringToFront: true);
                }
            }
            catch (Exception ex)
            {
                _logger.Error("ShowKernelWorkbookWindows failed.", ex);
            }
        }

        private bool ShouldPromoteKernelWorkbookOnHomeRelease()
        {
            Excel.Workbook activeWorkbook = null;
            try
            {
                activeWorkbook = _application.ActiveWorkbook;
            }
            catch (Exception ex)
            {
                _logger.Error("ShouldPromoteKernelWorkbookOnHomeRelease failed to resolve ActiveWorkbook.", ex);
            }

            bool hasActiveWorkbook = activeWorkbook != null;
            bool isActiveWorkbookKernel = hasActiveWorkbook && IsKernelWorkbook(activeWorkbook);
            bool hasVisibleNonKernelWorkbook = !hasActiveWorkbook && HasVisibleNonKernelWorkbook();
            bool shouldPromoteKernelWorkbook = KernelWorkbookPromotionPolicy.ShouldPromoteKernelWorkbookOnHomeRelease(
                isKernelCaseCreationFlowActive: _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                hasActiveWorkbook: hasActiveWorkbook,
                isActiveWorkbookKernel: isActiveWorkbookKernel,
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook);

            if (shouldPromoteKernelWorkbook)
            {
                return true;
            }

            _logger.Info(
                "Kernel workbook promotion skipped to preserve active non-kernel workbook. activeWorkbook="
                + _excelInteropService.GetWorkbookFullName(activeWorkbook));
            return false;
        }

        private void EnsureExcelApplicationVisible()
        {
            try
            {
                _application.Visible = true;
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                ShowWindow(hwnd, SwRestore);
                ShowWindow(hwnd, SwShow);
            }
            catch
            {
            }
        }

        private void ShowExcelMainWindow()
        {
            try
            {
                EnsureExcelApplicationVisible();
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                SetForegroundWindow(hwnd);
            }
            catch
            {
            }
        }

        private void ConcealKernelWorkbookWindowsForHomeDisplay(Excel.Workbook workbook, string triggerReason)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        if (window == null)
                        {
                            continue;
                        }

                        SetKernelWindowVisibleFalse(
                            workbook,
                            window,
                            index,
                            "ConcealKernelWorkbookWindowsForHomeDisplay|" + (triggerReason ?? string.Empty));
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("ConcealKernelWorkbookWindowsForHomeDisplay window conceal failed. index=" + index.ToString(), ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("ConcealKernelWorkbookWindowsForHomeDisplay failed.", ex);
            }
        }

        private void SetKernelWindowVisibleFalse(Excel.Workbook workbook, Excel.Window window, int index, string triggerReason)
        {
            string caller = ResolveExternalCaller();
            string beforeState = FormatWindowDescriptor(window);
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=set-window-visible-false-start trigger="
                + (triggerReason ?? string.Empty)
                + ", caller="
                + caller
                + ", index="
                + index.ToString()
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", windowBefore="
                + beforeState
                + ", activeState="
                + FormatActiveExcelState());

            try
            {
                if (window == null)
                {
                    LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=set-window-visible-false-end trigger="
                        + (triggerReason ?? string.Empty)
                        + ", caller="
                        + caller
                        + ", index="
                        + index.ToString()
                        + ", result=skipped-null-window, workbook="
                        + FormatWorkbookDescriptor(workbook));
                    return;
                }

                window.Visible = false;
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=set-window-visible-false-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", caller="
                    + caller
                    + ", index="
                    + index.ToString()
                    + ", result=success, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", windowBefore="
                    + beforeState
                    + ", windowAfter="
                    + FormatWindowDescriptor(window));
            }
            catch (Exception ex)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=set-window-visible-false-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", caller="
                    + caller
                    + ", index="
                    + index.ToString()
                    + ", result=failed, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", window="
                    + FormatWindowDescriptor(window)
                    + ", exceptionType="
                    + ex.GetType().Name
                    + ", exceptionMessage="
                    + (ex.Message ?? string.Empty));
                throw;
            }
        }

        private void LogKernelFlickerTrace(string detail)
        {
            _logger.Info(KernelFlickerTracePrefix + " " + (detail ?? string.Empty));
        }

        private string FormatActiveExcelState()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            Excel.Window activeWindow = _excelInteropService == null ? null : _excelInteropService.GetActiveWindow();
            return "activeWorkbook=" + FormatWorkbookDescriptor(activeWorkbook) + ",activeWindow=" + FormatWindowDescriptor(activeWindow);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\",visible=\""
                + SafeWindowVisible(window)
                + "\",state=\""
                + SafeWindowState(window)
                + "\"";
        }

        private string DescribeVisibleNonKernelWorkbookWindows()
        {
            if (_application == null)
            {
                return "app-null";
            }

            List<string> descriptors = new List<string>();
            try
            {
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (workbook == null || IsKernelWorkbook(workbook))
                    {
                        continue;
                    }

                    List<string> visibleWindows = new List<string>();
                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (SafeWindowVisibleValue(window))
                        {
                            visibleWindows.Add(FormatWindowDescriptor(window));
                        }
                    }

                    if (visibleWindows.Count > 0)
                    {
                        descriptors.Add(FormatWorkbookDescriptor(workbook) + " windows=[" + string.Join(" | ", visibleWindows) + "]");
                    }
                }
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }

            return descriptors.Count == 0 ? "none" : string.Join(" || ", descriptors);
        }

        private string DescribeVisibleOtherWorkbookWindows(Excel.Workbook workbookToIgnore)
        {
            if (_application == null)
            {
                return "app-null";
            }

            List<string> descriptors = new List<string>();
            try
            {
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (workbook == null || ReferenceEquals(workbook, workbookToIgnore))
                    {
                        continue;
                    }

                    List<string> visibleWindows = new List<string>();
                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (SafeWindowVisibleValue(window))
                        {
                            visibleWindows.Add(FormatWindowDescriptor(window));
                        }
                    }

                    if (visibleWindows.Count > 0)
                    {
                        descriptors.Add(FormatWorkbookDescriptor(workbook) + " windows=[" + string.Join(" | ", visibleWindows) + "]");
                    }
                }
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }

            return descriptors.Count == 0 ? "none" : string.Join(" || ", descriptors);
        }

        private string DescribeWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return "none";
            }

            List<string> descriptors = new List<string>();
            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = workbook.Windows[index];
                    descriptors.Add("index=" + index.ToString() + "," + FormatWindowDescriptor(window));
                }
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }

            return descriptors.Count == 0 ? "none" : string.Join(" | ", descriptors);
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        private string SafeWorkbookName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookName(workbook);
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowCaption(Excel.Window window)
        {
            try
            {
                if (window == null)
                {
                    return string.Empty;
                }

                dynamic lateBoundWindow = window;
                return Convert.ToString(lateBoundWindow.Caption) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowVisible(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.Visible.ToString();
            }
            catch
            {
                return "error";
            }
        }

        private static bool SafeWindowVisibleValue(Excel.Window window)
        {
            try
            {
                return window != null && window.Visible;
            }
            catch
            {
                return false;
            }
        }

        private static string SafeWindowState(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.WindowState.ToString();
            }
            catch
            {
                return "error";
            }
        }

        private static string ResolveExternalCaller()
        {
            try
            {
                StackTrace stackTrace = new StackTrace(skipFrames: 1, fNeedFileInfo: false);
                StackFrame[] frames = stackTrace.GetFrames();
                if (frames == null)
                {
                    return string.Empty;
                }

                foreach (StackFrame frame in frames)
                {
                    var method = frame.GetMethod();
                    if (method == null)
                    {
                        continue;
                    }

                    Type declaringType = method.DeclaringType;
                    if (declaringType == typeof(KernelWorkbookService))
                    {
                        continue;
                    }

                    string typeName = declaringType == null ? string.Empty : declaringType.FullName ?? string.Empty;
                    return string.IsNullOrWhiteSpace(typeName) ? method.Name : typeName + "." + method.Name;
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private sealed class KernelHomeSessionCloseCoordinator
        {
            private readonly KernelWorkbookService _owner;

            internal KernelHomeSessionCloseCoordinator(KernelWorkbookService owner)
            {
                _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            }

            internal void Execute(bool saveKernelWorkbook, string entryPoint, string caller)
            {
                Excel.Workbook workbook = _owner.ResolveWorkbookForHomeDisplayOrClose("CloseHomeSession");
                bool otherVisibleWorkbookExists = _owner.HasOtherVisibleWorkbookCore(workbook);
                bool otherWorkbookExists = _owner.HasOtherWorkbookCore(workbook);
                bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                    saveKernelWorkbook,
                    _owner._kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                    otherVisibleWorkbookExists,
                    otherWorkbookExists);
                KernelHomeSessionCompletionAction completionAction = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                    skipDisplayRestoreForCaseCreation,
                    otherVisibleWorkbookExists,
                    otherWorkbookExists);
                _owner.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-enter entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", caller="
                    + caller
                    + ", saveKernelWorkbook="
                    + saveKernelWorkbook.ToString()
                    + ", workbook="
                    + _owner.FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _owner.FormatActiveExcelState()
                    + ", otherVisibleWorkbookExists="
                    + otherVisibleWorkbookExists.ToString()
                    + ", otherWorkbookExists="
                    + otherWorkbookExists.ToString()
                    + ", skipDisplayRestoreForCaseCreation="
                    + skipDisplayRestoreForCaseCreation.ToString()
                    + ", completionAction="
                    + completionAction.ToString()
                    + ", otherVisibleTargets="
                    + _owner.DescribeVisibleOtherWorkbookWindows(workbook));
                _owner._logger.Info(
                    "CloseHomeSession started. saveKernelWorkbook="
                    + saveKernelWorkbook.ToString()
                    + ", workbook="
                    + _owner.GetWorkbookFullNameCore(workbook)
                    + ", otherVisibleWorkbookExists="
                    + otherVisibleWorkbookExists.ToString()
                    + ", otherWorkbookExists="
                    + otherWorkbookExists.ToString()
                    + ", skipDisplayRestoreForCaseCreation="
                    + skipDisplayRestoreForCaseCreation.ToString());

                if (workbook != null && !ExecuteCloseBranch(workbook, saveKernelWorkbook, skipDisplayRestoreForCaseCreation, entryPoint))
                {
                    return;
                }

                CompleteHomeSession(saveKernelWorkbook, completionAction, workbook, entryPoint);
            }

            private bool ExecuteCloseBranch(
                Excel.Workbook workbook,
                bool saveKernelWorkbook,
                bool skipDisplayRestoreForCaseCreation,
                string entryPoint)
            {
                if (saveKernelWorkbook)
                {
                    _owner.LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=save-and-close, workbook="
                        + _owner.FormatWorkbookDescriptor(workbook)
                        + ", skipDisplayRestoreForCaseCreation="
                        + skipDisplayRestoreForCaseCreation.ToString());
                    if (skipDisplayRestoreForCaseCreation)
                    {
                        _owner.ConcealKernelWorkbookWindowsForCaseCreationCloseCore(workbook);
                    }

                    _owner.SaveAndCloseKernelWorkbookCore(workbook);
                    return true;
                }

                if (_owner._kernelWorkbookLifecycleService != null)
                {
                    _owner.LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=request-managed-close, workbook="
                        + _owner.FormatWorkbookDescriptor(workbook)
                        + ", lifecycleAvailable=True, activeState="
                        + _owner.FormatActiveExcelState());
                    bool closeScheduled = _owner.RequestManagedCloseFromHomeExitCore(workbook);
                    _owner.LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch-result entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=request-managed-close, workbook="
                        + _owner.FormatWorkbookDescriptor(workbook)
                        + ", closeScheduled="
                        + closeScheduled.ToString());
                    if (!closeScheduled)
                    {
                        _owner.LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=close-home-session-end entryPoint="
                            + (entryPoint ?? string.Empty)
                            + ", result=canceled-before-managed-close, workbook="
                            + _owner.FormatWorkbookDescriptor(workbook));
                        _owner._logger.Info("CloseHomeSession canceled before managed close was scheduled.");
                        return false;
                    }

                    return true;
                }

                _owner.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", branch=close-without-lifecycle, workbook="
                    + _owner.FormatWorkbookDescriptor(workbook)
                    + ", lifecycleAvailable=False");
                _owner.CloseKernelWorkbookWithoutLifecycleCore(workbook);
                return true;
            }

            private void CompleteHomeSession(
                bool saveKernelWorkbook,
                KernelHomeSessionCompletionAction completionAction,
                Excel.Workbook workbook,
                string entryPoint)
            {
                _owner.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-completion entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", completionAction="
                    + completionAction.ToString()
                    + ", workbook="
                    + _owner.FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _owner.FormatActiveExcelState());
                if (completionAction == KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithoutShowingExcelAndQuit)
                {
                    _owner.ReleaseHomeDisplayCore(false);
                    if (saveKernelWorkbook || _owner._kernelWorkbookLifecycleService == null)
                    {
                        _owner.QuitApplicationCore();
                    }
                }
                else if (completionAction == KernelHomeSessionCompletionAction.DismissPreparedHomeDisplayState)
                {
                    _owner.DismissPreparedHomeDisplayStateCore("CloseHomeSession.CaseCreationSkipRestore");
                }
                else
                {
                    _owner.ReleaseHomeDisplayCore(true);
                }

                _owner.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-end entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", result=completed, workbook="
                    + _owner.FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _owner.FormatActiveExcelState());
                _owner.ClearHomeWorkbookBinding("CloseHomeSession.Completed");
                _owner._logger.Info("CloseHomeSession completed. saveKernelWorkbook=" + saveKernelWorkbook.ToString());
            }
        }

        private sealed class KernelHomeBinding
        {
            internal KernelHomeBinding(Excel.Workbook workbook, string systemRoot)
            {
                Workbook = workbook;
                SystemRoot = systemRoot ?? string.Empty;
            }

            internal Excel.Workbook Workbook { get; }

            internal string SystemRoot { get; }
        }

        private enum KernelHomeBindingStatus
        {
            None = 0,
            Valid = 1,
            Invalid = 2
        }

        internal sealed class KernelWorkbookServiceTestHooks
        {
            internal Action ApplyHomeDisplayVisibility { get; set; }

            internal Func<Excel.Workbook> GetOpenKernelWorkbook { get; set; }

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

    internal sealed class KernelSettingsState
    {
        internal string SystemRoot { get; set; } = string.Empty;
        internal string DefaultRoot { get; set; } = string.Empty;
        internal string NameRuleA { get; set; } = "YYYY";
        internal string NameRuleB { get; set; } = "DOC";
    }
}





