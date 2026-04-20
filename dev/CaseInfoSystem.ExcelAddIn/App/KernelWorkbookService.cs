using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private const string SystemRootPropName = "SYSTEM_ROOT";
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
        private KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private bool _isHomeDisplayPrepared;

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
        }

        internal Excel.Workbook GetOpenKernelWorkbook()
        {
            try
            {
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (WorkbookFileNameResolver.IsKernelWorkbookName(_excelInteropService.GetWorkbookName(workbook)))
                    {
                        return workbook;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("GetOpenKernelWorkbook failed.", ex);
            }

            return null;
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return WorkbookFileNameResolver.IsKernelWorkbookName(_excelInteropService.GetWorkbookName(workbook));
        }

        internal Excel.Workbook ResolveKernelWorkbook(Domain.WorkbookContext context)
        {
            Excel.Workbook openKernelWorkbook = GetOpenKernelWorkbookCore();
            string kernelPath = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: openKernelWorkbook != null,
                systemRoot: context == null ? string.Empty : context.SystemRoot,
                resolvePath: root => ResolveKernelWorkbookPathCore(root));

            if (openKernelWorkbook != null)
            {
                return openKernelWorkbook;
            }

            if (string.IsNullOrWhiteSpace(kernelPath))
            {
                return null;
            }

            return FindOpenWorkbookCore(kernelPath);
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
            bool hasExplicitKernelStartupContext = HasExplicitKernelStartupContext(startupWorkbook);
            bool hasKernelWorkbookContext = hasExplicitKernelStartupContext && HasKernelWorkbookContext();
            bool isStartupWorkbookKernel = hasExplicitKernelStartupContext && hasKernelWorkbookContext && IsKernelWorkbook(startupWorkbook);
            bool hasVisibleNonKernelWorkbook = hasExplicitKernelStartupContext && hasKernelWorkbookContext && !isStartupWorkbookKernel && HasVisibleNonKernelWorkbook();
            return KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext,
                hasKernelWorkbookContext,
                isStartupWorkbookKernel,
                hasVisibleNonKernelWorkbook);
        }

        internal string DescribeStartupState()
        {
            string activeWorkbookName = "(null)";
            bool activeIsKernel = false;
            bool hasOpenKernelWorkbook = false;
            bool hasVisibleNonKernelWorkbook = false;

            try
            {
                Excel.Workbook activeWorkbook = _application.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    activeWorkbookName = _excelInteropService.GetWorkbookName(activeWorkbook);
                    activeIsKernel = IsKernelWorkbook(activeWorkbook);
                }
            }
            catch
            {
                activeWorkbookName = "(error)";
            }

            try
            {
                hasOpenKernelWorkbook = GetOpenKernelWorkbook() != null;
                hasVisibleNonKernelWorkbook = HasVisibleNonKernelWorkbook();
            }
            catch
            {
            }

            return "activeWorkbook="
                + activeWorkbookName
                + ", activeIsKernel="
                + activeIsKernel
                + ", hasOpenKernelWorkbook="
                + hasOpenKernelWorkbook
                + ", hasVisibleNonKernelWorkbook="
                + hasVisibleNonKernelWorkbook;
        }

        internal KernelSettingsState LoadSettings()
        {
            Excel.Workbook workbook = GetOrOpenKernelWorkbook();
            if (workbook == null)
            {
                _logger.Warn("Kernel settings load skipped because kernel workbook was not available.");
                return new KernelSettingsState();
            }

            string nameRuleA = _excelInteropService.TryGetDocumentProperty(workbook, "NAME_RULE_A");
            string nameRuleB = _excelInteropService.TryGetDocumentProperty(workbook, "NAME_RULE_B");
            string defaultRoot = _excelInteropService.TryGetDocumentProperty(workbook, "DEFAULT_ROOT");
            string systemRoot = TryGetKernelWorkbookDirectory(workbook);

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
                NameRuleA = KernelNamingService.NormalizeNameRuleA(string.IsNullOrWhiteSpace(nameRuleA) ? "YYYY" : nameRuleA),
                NameRuleB = KernelNamingService.NormalizeNameRuleB(string.IsNullOrWhiteSpace(nameRuleB) ? "DOC" : nameRuleB)
            };
        }

        internal string ResolveCurrentCaseWorkbookExtension(string systemRoot)
        {
            string baseWorkbookPath = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath(systemRoot, _pathCompatibilityService);
            return WorkbookFileNameResolver.GetWorkbookExtensionOrDefault(baseWorkbookPath);
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

        internal void SaveNameRuleA(string ruleA)
        {
            Excel.Workbook workbook = GetOrOpenKernelWorkbook();
            if (workbook == null)
            {
                return;
            }

            SetDocumentProperty(workbook, "NAME_RULE_A", ruleA);
            workbook.Save();
        }

        internal void SaveNameRuleB(string ruleB)
        {
            Excel.Workbook workbook = GetOrOpenKernelWorkbook();
            if (workbook == null)
            {
                return;
            }

            SetDocumentProperty(workbook, "NAME_RULE_B", ruleB);
            workbook.Save();
        }

        internal string SelectAndSaveDefaultRoot()
        {
            Excel.Workbook workbook = GetOrOpenKernelWorkbook();
            if (workbook == null)
            {
                return null;
            }

            string selectedPath = SelectFolderPath("既定フォルダを選択してください。", _excelInteropService.TryGetDocumentProperty(workbook, "DEFAULT_ROOT"));
            if (string.IsNullOrWhiteSpace(selectedPath))
            {
                return null;
            }

            SetDocumentProperty(workbook, "DEFAULT_ROOT", selectedPath);
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
            Excel.Workbook workbook = GetOpenKernelWorkbookCore();
            bool otherVisibleWorkbookExists = HasOtherVisibleWorkbookCore(workbook);
            bool otherWorkbookExists = HasOtherWorkbookCore(workbook);
            bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                saveKernelWorkbook,
                _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                otherVisibleWorkbookExists,
                otherWorkbookExists);
            KernelHomeSessionCompletionAction completionAction = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                skipDisplayRestoreForCaseCreation,
                otherVisibleWorkbookExists,
                otherWorkbookExists);
            string caller = ResolveExternalCaller();
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=close-home-session-enter entryPoint="
                + (entryPoint ?? string.Empty)
                + ", caller="
                + caller
                + ", saveKernelWorkbook="
                + saveKernelWorkbook.ToString()
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState()
                + ", otherVisibleWorkbookExists="
                + otherVisibleWorkbookExists.ToString()
                + ", otherWorkbookExists="
                + otherWorkbookExists.ToString()
                + ", skipDisplayRestoreForCaseCreation="
                + skipDisplayRestoreForCaseCreation.ToString()
                + ", completionAction="
                + completionAction.ToString()
                + ", otherVisibleTargets="
                + DescribeVisibleOtherWorkbookWindows(workbook));
            _logger.Info(
                "CloseHomeSession started. saveKernelWorkbook="
                + saveKernelWorkbook.ToString()
                + ", workbook="
                + GetWorkbookFullNameCore(workbook)
                + ", otherVisibleWorkbookExists="
                + otherVisibleWorkbookExists.ToString()
                + ", otherWorkbookExists="
                + otherWorkbookExists.ToString()
                + ", skipDisplayRestoreForCaseCreation="
                + skipDisplayRestoreForCaseCreation.ToString());

            if (workbook != null)
            {
                if (saveKernelWorkbook)
                {
                    LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=save-and-close, workbook="
                        + FormatWorkbookDescriptor(workbook)
                        + ", skipDisplayRestoreForCaseCreation="
                        + skipDisplayRestoreForCaseCreation.ToString());
                    // 処理ブロック: CASE 作成完了直後は Kernel シートが前景に出ないよう、閉じる直前に window を不可視化する。
                    if (skipDisplayRestoreForCaseCreation)
                    {
                        ConcealKernelWorkbookWindowsForCaseCreationCloseCore(workbook);
                    }

                    SaveAndCloseKernelWorkbookCore(workbook);
                }
                else if (_kernelWorkbookLifecycleService != null)
                {
                    LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=request-managed-close, workbook="
                        + FormatWorkbookDescriptor(workbook)
                        + ", lifecycleAvailable=True, activeState="
                        + FormatActiveExcelState());
                    bool closeScheduled = RequestManagedCloseFromHomeExitCore(workbook);
                    LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch-result entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=request-managed-close, workbook="
                        + FormatWorkbookDescriptor(workbook)
                        + ", closeScheduled="
                        + closeScheduled.ToString());
                    if (!closeScheduled)
                    {
                        LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=close-home-session-end entryPoint="
                            + (entryPoint ?? string.Empty)
                            + ", result=canceled-before-managed-close, workbook="
                            + FormatWorkbookDescriptor(workbook));
                        _logger.Info("CloseHomeSession canceled before managed close was scheduled.");
                        return;
                    }
                }
                else
                {
                    LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=close-without-lifecycle, workbook="
                        + FormatWorkbookDescriptor(workbook)
                        + ", lifecycleAvailable=False");
                    CloseKernelWorkbookWithoutLifecycleCore(workbook);
                }
            }

            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=close-home-session-completion entryPoint="
                + (entryPoint ?? string.Empty)
                + ", completionAction="
                + completionAction.ToString()
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            if (completionAction == KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithoutShowingExcelAndQuit)
            {
                ReleaseHomeDisplayCore(false);
                if (saveKernelWorkbook || _kernelWorkbookLifecycleService == null)
                {
                    QuitApplicationCore();
                }
            }
            else if (completionAction == KernelHomeSessionCompletionAction.DismissPreparedHomeDisplayState)
            {
                // 処理ブロック: CASE 作成完了後は、既に表示中の CASE などの前景を維持し、Kernel 復帰を行わない。
                DismissPreparedHomeDisplayStateCore("CloseHomeSession.CaseCreationSkipRestore");
            }
            else
            {
                ReleaseHomeDisplayCore(true);
            }

            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=close-home-session-end entryPoint="
                + (entryPoint ?? string.Empty)
                + ", result=completed, workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger.Info("CloseHomeSession completed. saveKernelWorkbook=" + saveKernelWorkbook.ToString());
        }

        private Excel.Workbook GetOpenKernelWorkbookCore()
        {
            return _testHooks != null && _testHooks.GetOpenKernelWorkbook != null
                ? _testHooks.GetOpenKernelWorkbook()
                : GetOpenKernelWorkbook();
        }

        private string ResolveKernelWorkbookPathCore(string systemRoot)
        {
            return _testHooks != null && _testHooks.ResolveKernelWorkbookPath != null
                ? _testHooks.ResolveKernelWorkbookPath(systemRoot)
                : WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private Excel.Workbook FindOpenWorkbookCore(string workbookPath)
        {
            return _testHooks != null && _testHooks.FindOpenWorkbook != null
                ? _testHooks.FindOpenWorkbook(workbookPath)
                : _excelInteropService.FindOpenWorkbook(workbookPath);
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
            Excel.Workbook workbook = GetOrOpenKernelWorkbook();
            if (workbook == null)
            {
                return false;
            }

            bool previousEnableEvents = _application.EnableEvents;
            try
            {
                _application.EnableEvents = false;
                ReleaseHomeDisplay(true);
                EnsureWorkbookVisible(workbook);
                Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName(workbook, codeName);
                if (worksheet == null)
                {
                    return false;
                }

                workbook.Activate();
                worksheet.Activate();
                return true;
            }
            finally
            {
                _application.EnableEvents = previousEnableEvents;
            }
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

            string workbookPath = ResolveKernelWorkbookPathFromAvailableSystemRoot();
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

        private string ResolveKernelWorkbookPathFromAvailableSystemRoot()
        {
            string systemRoot = ResolveSystemRootFromAvailableWorkbooks();
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                return null;
            }

            return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private string ResolveSystemRootFromAvailableWorkbooks()
        {
            string systemRoot = GetSystemRootFromWorkbook(_application.ActiveWorkbook);
            if (!string.IsNullOrWhiteSpace(systemRoot))
            {
                return systemRoot;
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                systemRoot = GetSystemRootFromWorkbook(workbook);
                if (!string.IsNullOrWhiteSpace(systemRoot))
                {
                    return systemRoot;
                }
            }

            return string.Empty;
        }

        private string GetSystemRootFromWorkbook(Excel.Workbook workbook)
        {
            string systemRoot = NormalizePath(_excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropName));
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                string workbookPathFallback = TryGetKernelWorkbookDirectory(workbook);
                if (string.IsNullOrWhiteSpace(workbookPathFallback))
                {
                    return string.Empty;
                }

                _logger.Info("GetSystemRootFromWorkbook fallback used workbook directory. workbook=" + _excelInteropService.GetWorkbookFullName(workbook));
                return workbookPathFallback;
            }

            string kernelWorkbookPath = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
            return string.IsNullOrWhiteSpace(kernelWorkbookPath) ? string.Empty : systemRoot;
        }

        private string TryGetKernelWorkbookDirectory(Excel.Workbook workbook)
        {
            if (!IsKernelWorkbook(workbook))
            {
                return string.Empty;
            }

            string workbookFullName = NormalizePath(_excelInteropService.GetWorkbookFullName(workbook));
            if (string.IsNullOrWhiteSpace(workbookFullName) || !File.Exists(workbookFullName))
            {
                return string.Empty;
            }

            string workbookDirectory;
            try
            {
                workbookDirectory = Path.GetDirectoryName(workbookFullName);
            }
            catch (Exception ex)
            {
                _logger.Error("TryGetKernelWorkbookDirectory Path.GetDirectoryName failed.", ex);
                return string.Empty;
            }

            workbookDirectory = NormalizePath(workbookDirectory);
            if (string.IsNullOrWhiteSpace(workbookDirectory))
            {
                return string.Empty;
            }

            string kernelWorkbookPath = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(workbookDirectory, _pathCompatibilityService);
            return string.IsNullOrWhiteSpace(kernelWorkbookPath) ? string.Empty : workbookDirectory;
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar);
            }
            catch
            {
                return path.Trim();
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

        private string SelectFolderPath(string dialogTitle, string initialDirectory)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Multiselect = false;
                dialog.Title = dialogTitle;
                dialog.EnsurePathExists = true;
                dialog.AllowNonFileSystemItems = false;

                if (!string.IsNullOrWhiteSpace(initialDirectory) && Directory.Exists(initialDirectory))
                {
                    dialog.InitialDirectory = initialDirectory;
                    dialog.DefaultDirectory = initialDirectory;
                }

                if (dialog.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    return null;
                }

                return dialog.FileName;
            }
        }

        private static void SetDocumentProperty(Excel.Workbook workbook, string propertyName, string value)
        {
            dynamic properties = workbook.CustomDocumentProperties;
            try
            {
                properties[propertyName].Value = value;
            }
            catch
            {
                const int MsoPropertyTypeString = 4;
                properties.Add(propertyName, false, MsoPropertyTypeString, value);
            }
        }

        private bool HasOtherVisibleWorkbook(Excel.Workbook workbookToIgnore)
        {
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                {
                    continue;
                }

                if (workbook.Windows.Count > 0 && workbook.Windows.Cast<Excel.Window>().Any(window => window.Visible))
                {
                    return true;
                }
            }

            return false;
        }

        private bool HasOtherWorkbook(Excel.Workbook workbookToIgnore)
        {
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        private bool HasVisibleNonKernelWorkbook()
        {
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (IsKernelWorkbook(workbook))
                {
                    continue;
                }

                if (workbook.Windows.Count > 0 && workbook.Windows.Cast<Excel.Window>().Any(window => window.Visible))
                {
                    return true;
                }
            }

            return false;
        }

        private bool HasExplicitKernelStartupContext(Excel.Workbook startupWorkbook)
        {
            if (IsKernelWorkbook(startupWorkbook))
            {
                return true;
            }

            try
            {
                return IsKernelWorkbook(_application.ActiveWorkbook);
            }
            catch
            {
                return false;
            }
        }

        private bool HasKernelWorkbookContext()
        {
            try
            {
                if (IsKernelWorkbook(_application.ActiveWorkbook))
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                return GetOpenKernelWorkbook() != null;
            }
            catch
            {
                return false;
            }
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
            Excel.Workbook kernelWorkbook = GetOpenKernelWorkbookCore();
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
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                ShowWindow(hwnd, SwHide);
                _application.Visible = false;
            }
            catch
            {
            }
        }

        private void HideKernelWorkbookWindows()
        {
            HideKernelWorkbookWindows("HideKernelWorkbookWindows.ResolveOpenKernelWorkbook", GetOpenKernelWorkbook());
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

        private void BringExcelToFront()
        {
            try
            {
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                ShowWindow(hwnd, SwRestore);
                SetForegroundWindow(hwnd);
            }
            catch
            {
            }
        }

        private void BringWorkbookWindowToFront(Excel.Window window)
        {
            if (window == null)
            {
                return;
            }

            try
            {
                IntPtr hwnd = new IntPtr(window.Hwnd);
                ShowWindow(hwnd, SwRestore);
                SetForegroundWindow(hwnd);
            }
            catch
            {
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





