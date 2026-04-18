using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookService
    {
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
            Excel.Workbook openKernelWorkbook = GetOpenKernelWorkbook();
            string kernelPath = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: openKernelWorkbook != null,
                systemRoot: context == null ? string.Empty : context.SystemRoot,
                resolvePath: root => WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(root, _pathCompatibilityService));

            if (openKernelWorkbook != null)
            {
                return openKernelWorkbook;
            }

            if (string.IsNullOrWhiteSpace(kernelPath))
            {
                return null;
            }

            return _excelInteropService.FindOpenWorkbook(kernelPath);
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

            ApplyHomeDisplayVisibility();
            _isHomeDisplayPrepared = true;
        }

        internal void PrepareForHomeDisplayFromSheet()
        {
            ApplyHomeDisplayVisibility();
            _isHomeDisplayPrepared = true;
        }

        internal void CompleteHomeNavigation(bool showExcel)
        {
            ReleaseHomeDisplay(showExcel);
        }

        internal void EnsureHomeDisplayHidden()
        {
            if (!_isHomeDisplayPrepared)
            {
                return;
            }

            ApplyHomeDisplayVisibility();
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
            CloseHomeSession(saveKernelWorkbook: false);
        }

        internal void CloseHomeSessionSavingKernel()
        {
            CloseHomeSession(saveKernelWorkbook: true);
        }

        private void CloseHomeSession(bool saveKernelWorkbook)
        {
            Excel.Workbook workbook = GetOpenKernelWorkbook();
            bool otherVisibleWorkbookExists = HasOtherVisibleWorkbook(workbook);
            bool otherWorkbookExists = HasOtherWorkbook(workbook);
            bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                saveKernelWorkbook,
                _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                otherVisibleWorkbookExists,
                otherWorkbookExists);
            KernelHomeSessionCompletionAction completionAction = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                skipDisplayRestoreForCaseCreation,
                otherVisibleWorkbookExists,
                otherWorkbookExists);
            _logger.Info(
                "CloseHomeSession started. saveKernelWorkbook="
                + saveKernelWorkbook.ToString()
                + ", workbook="
                + (workbook == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook))
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
                    // 処理ブロック: CASE 作成完了直後は Kernel シートが前景に出ないよう、閉じる直前に window を不可視化する。
                    if (skipDisplayRestoreForCaseCreation)
                    {
                        ConcealKernelWorkbookWindowsForCaseCreationClose(workbook);
                    }

                    SaveAndCloseKernelWorkbook(workbook);
                }
                else if (_kernelWorkbookLifecycleService != null)
                {
                    bool closeScheduled = _kernelWorkbookLifecycleService.RequestManagedCloseFromHomeExit(workbook);
                    if (!closeScheduled)
                    {
                        _logger.Info("CloseHomeSession canceled before managed close was scheduled.");
                        return;
                    }
                }
                else
                {
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
            }

            if (completionAction == KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithoutShowingExcelAndQuit)
            {
                ReleaseHomeDisplay(false);
                if (saveKernelWorkbook || _kernelWorkbookLifecycleService == null)
                {
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
            }
            else if (completionAction == KernelHomeSessionCompletionAction.DismissPreparedHomeDisplayState)
            {
                // 処理ブロック: CASE 作成完了後は、既に表示中の CASE などの前景を維持し、Kernel 復帰を行わない。
                DismissPreparedHomeDisplayState("CloseHomeSession.CaseCreationSkipRestore");
            }
            else
            {
                ReleaseHomeDisplay(true);
            }

            _logger.Info("CloseHomeSession completed. saveKernelWorkbook=" + saveKernelWorkbook.ToString());
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
                    workbook.Windows[1].Visible = false;
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
                    workbook.Saved = true;
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

                        window.Visible = false;
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

        private void ApplyHomeDisplayVisibility()
        {
            if (HasVisibleNonKernelWorkbook())
            {
                if (!ShouldPreserveOtherWorkbookWindowLayout())
                {
                    EnsureExcelApplicationVisible();
                }

                HideKernelWorkbookWindows();
                _logger.Info(
                    "ApplyHomeDisplayVisibility minimized kernel windows because a non-kernel workbook is visible. preserveOtherWindowLayout="
                    + ShouldPreserveOtherWorkbookWindowLayout().ToString());
                return;
            }

            HideExcelMainWindow();
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
            HideKernelWorkbookWindows(GetOpenKernelWorkbook());
        }

        private void HideKernelWorkbookWindows(Excel.Workbook workbook)
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
                        if (window != null)
                        {
                            window.WindowState = Excel.XlWindowState.xlMinimized;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("HideKernelWorkbookWindows window minimize failed. index=" + index.ToString(), ex);
                    }
                }
            }
            catch (Exception ex)
            {
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
    }

    internal sealed class KernelSettingsState
    {
        internal string SystemRoot { get; set; } = string.Empty;
        internal string DefaultRoot { get; set; } = string.Empty;
        internal string NameRuleA { get; set; } = "YYYY";
        internal string NameRuleB { get; set; } = "DOC";
    }
}





