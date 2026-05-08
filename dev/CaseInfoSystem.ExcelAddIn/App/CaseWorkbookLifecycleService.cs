using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // クラス: CASE/Base ブックの保存・終了まわりのライフサイクルを管理する。
    // 責務: 初期化、終了確認、管理対象クローズ、後追い終了処理を制御する。
    internal sealed class CaseWorkbookLifecycleService
    {
        private const string CaseHomeSheetCodeName = "shHOME";
        private const string NameRuleAPropertyName = "NAME_RULE_A";
        private const string NameRuleBPropertyName = "NAME_RULE_B";

        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly TransientPaneSuppressionService _transientPaneSuppressionService;
        private readonly ManagedCloseState _managedCloseState;
        private readonly CaseClosePromptService _caseClosePromptService;
        private readonly CaseFolderOpenService _caseFolderOpenService;
        private readonly KernelNameRuleReader _kernelNameRuleReader;
        private readonly PostCloseFollowUpScheduler _postCloseFollowUpScheduler;
        private readonly Logger _logger;
        private readonly HashSet<string> _initializedWorkbookKeys;
        private readonly HashSet<string> _sessionDirtyWorkbookKeys;
        private readonly HashSet<string> _createdCaseFolderOfferPendingWorkbookKeys;
        private readonly CaseWorkbookLifecycleServiceTestHooks _testHooks;
        private Control _managedCloseDispatcher;

        internal CaseWorkbookLifecycleService(
            WorkbookRoleResolver workbookRoleResolver,
            Excel.Application application,
            ExcelInteropService excelInteropService,
            TransientPaneSuppressionService transientPaneSuppressionService,
            ManagedCloseState managedCloseState,
            CaseClosePromptService caseClosePromptService,
            CaseFolderOpenService caseFolderOpenService,
            KernelNameRuleReader kernelNameRuleReader,
            PostCloseFollowUpScheduler postCloseFollowUpScheduler,
            Logger logger)
        {
            _workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException(nameof(workbookRoleResolver));
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException(nameof(transientPaneSuppressionService));
            _managedCloseState = managedCloseState ?? throw new ArgumentNullException(nameof(managedCloseState));
            _caseClosePromptService = caseClosePromptService ?? throw new ArgumentNullException(nameof(caseClosePromptService));
            _caseFolderOpenService = caseFolderOpenService ?? throw new ArgumentNullException(nameof(caseFolderOpenService));
            _kernelNameRuleReader = kernelNameRuleReader ?? throw new ArgumentNullException(nameof(kernelNameRuleReader));
            _postCloseFollowUpScheduler = postCloseFollowUpScheduler ?? throw new ArgumentNullException(nameof(postCloseFollowUpScheduler));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _initializedWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _sessionDirtyWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _createdCaseFolderOfferPendingWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _testHooks = null;
        }

        internal CaseWorkbookLifecycleService(Logger logger, CaseWorkbookLifecycleServiceTestHooks testHooks)
        {
            _workbookRoleResolver = null;
            _application = null;
            _excelInteropService = null;
            _transientPaneSuppressionService = null;
            _managedCloseState = new ManagedCloseState();
            _caseClosePromptService = null;
            _caseFolderOpenService = null;
            _kernelNameRuleReader = null;
            _postCloseFollowUpScheduler = null;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _initializedWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _sessionDirtyWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _createdCaseFolderOfferPendingWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _testHooks = testHooks;
        }

        // メソッド: CASE/Base ブックの初回アクティブ化時に初期化処理を行う。
        // 引数: workbook - 対象ブック。
        // 戻り値: なし。
        // 副作用: Kernel からの設定同期や内部状態初期化を行う。
        internal void HandleWorkbookOpenedOrActivated(Excel.Workbook workbook)
        {
            string workbookKey = GetWorkbookKey(workbook);
            CaseWorkbookInitializationAction action = CaseWorkbookLifecycleInitializationPolicy.Decide(
                isBaseOrCaseWorkbook: IsBaseOrCaseWorkbookCore(workbook),
                workbookKey: workbookKey,
                isAlreadyInitialized: !string.IsNullOrWhiteSpace(workbookKey) && _initializedWorkbookKeys.Contains(workbookKey),
                isCaseWorkbook: IsCaseWorkbookCore(workbook));

            if (action == CaseWorkbookInitializationAction.None)
            {
                return;
            }

            try
            {
                if (action == CaseWorkbookInitializationAction.InitializeCaseWorkbook)
                {
                    RegisterKnownCaseWorkbookCore(workbook);
                    SyncNameRulesFromKernelToCaseCore(workbook);
                }

                _initializedWorkbookKeys.Add(workbookKey);
                _logger.Info("Case workbook lifecycle initialization completed. workbook=" + workbookKey);
            }
            catch (Exception ex)
            {
                // 例外処理: 初期化失敗で Excel の通常操作を止めないため、ログのみ残して継続する。
                _logger.Error("Case workbook lifecycle initialization failed.", ex);
            }
        }

        // メソッド: CASE のシート切り替え時の後続処理を扱う。
        // 引数: sheetObject - Excel が通知したアクティブシートオブジェクト。
        // 戻り値: なし。
        // 副作用: なし。
        internal void HandleSheetActivated(object sheetObject)
        {
            Excel.Worksheet worksheet = sheetObject as Excel.Worksheet;
            if (worksheet == null)
            {
                return;
            }

            Excel.Workbook workbook = worksheet.Parent as Excel.Workbook;
            if (!_workbookRoleResolver.IsCaseWorkbook(workbook))
            {
                return;
            }

            if (!IsCaseHomeSheet(worksheet))
            {
                return;
            }
        }

        // メソッド: CASE/Base ブック終了前の確認ダイアログと後続処理を制御する。
        // 引数: workbook - 対象ブック, cancel - Excel 終了を取り消すかどうか。
        // 戻り値: VSTO 側で終了処理を握った場合は true。
        // 副作用: メッセージ表示、管理クローズ予約、後追い終了予約を行う。
        internal bool HandleWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            bool isBaseOrCaseWorkbook = IsBaseOrCaseWorkbookCore(workbook);
            bool isManagedClose = IsManagedCloseCore(workbook);
            string workbookKey = GetWorkbookKey(workbook);
            bool isSessionDirty = _sessionDirtyWorkbookKeys.Contains(workbookKey);
            CaseWorkbookBeforeCloseAction action = CaseWorkbookBeforeClosePolicy.Decide(
                isBaseOrCaseWorkbook: isBaseOrCaseWorkbook,
                isManagedClose: isManagedClose,
                isSessionDirty: isSessionDirty);

            _logger.Info(
                "[KernelFlickerTrace] source=CaseWorkbookLifecycleService"
                + " action=workbook-close-immutable-facts-captured"
                + " workbook=" + (workbookKey ?? string.Empty)
                + ", isBaseOrCaseWorkbook=" + isBaseOrCaseWorkbook.ToString()
                + ", isManagedClose=" + isManagedClose.ToString()
                + ", isSessionDirty=" + isSessionDirty.ToString()
                + ", beforeCloseAction=" + action.ToString());

            if (action == CaseWorkbookBeforeCloseAction.Ignore)
            {
                return false;
            }

            if (action == CaseWorkbookBeforeCloseAction.SuppressPromptForManagedClose)
            {
                _logger.Info("Case workbook before-close prompt suppressed for managed close. workbook=" + workbookKey);
                return false;
            }

            try
            {
                string folderPath = ResolveContainingFolder(workbook);
                _logger.Info(
                    "[KernelFlickerTrace] source=CaseWorkbookLifecycleService"
                    + " action=workbook-close-follow-up-facts-captured"
                    + " workbook=" + (workbookKey ?? string.Empty)
                    + ", folderPathCaptured=" + (!string.IsNullOrWhiteSpace(folderPath)).ToString()
                    + ", beforeCloseAction=" + action.ToString());

                if (action == CaseWorkbookBeforeCloseAction.PromptForDirtySession)
                {
                    cancel = true;

                    DialogResult answer = ShowClosePromptCore(workbook);

                    if (answer == DialogResult.Cancel)
                    {
                        return true;
                    }

                    PromptToOpenCreatedCaseFolderIfNeeded(workbookKey, folderPath);
                    ScheduleManagedSessionCloseCore(workbookKey, folderPath, answer == DialogResult.Yes);
                    return true;
                }

                _logger.Info(
                    "Case workbook before-close follow-up is handled by VSTO. workbook="
                    + workbookKey
                    + ", macroEnabled="
                    + IsMacroEnabledWorkbook(workbook).ToString());

                PromptToOpenCreatedCaseFolderIfNeeded(workbookKey, folderPath);
                SchedulePostCloseFollowUpCore(workbookKey, folderPath);
                _logger.Info("Case workbook post-close follow-up posted. workbook=" + workbookKey);
                return false;
            }
            catch (Exception ex)
            {
                // 例外処理: 終了イベントで例外を再送出すると Excel のクローズ動作が不安定になるため、ログ化して返す。
                _logger.Error("Case workbook before-close handling failed.", ex);
                cancel = false;
                return false;
            }
        }

        // メソッド: VSTO 主導でブックを閉じる間だけ確認ダイアログを抑止するスコープを開始する。
        // 引数: workbook - 対象ブック。
        // 戻り値: スコープ解除用 IDisposable。
        // 副作用: 管理クローズ回数を更新する。
        internal IDisposable BeginManagedCloseScope(Excel.Workbook workbook)
        {
            string workbookKey = GetWorkbookKey(workbook);
            return _managedCloseState.BeginScope(workbookKey);
        }

        internal void RemoveWorkbookState(Excel.Workbook workbook)
        {
            string workbookKey = GetWorkbookKey(workbook);
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _initializedWorkbookKeys.Remove(workbookKey);
            _sessionDirtyWorkbookKeys.Remove(workbookKey);
            _createdCaseFolderOfferPendingWorkbookKeys.Remove(workbookKey);
            _managedCloseState.Remove(workbookKey);
            _workbookRoleResolver.RemoveKnownWorkbook(workbook);
        }

        internal void MarkCreatedCaseFolderOfferPending(Excel.Workbook workbook)
        {
            string workbookKey = GetWorkbookKey(workbook);
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _createdCaseFolderOfferPendingWorkbookKeys.Add(workbookKey);
            _logger.Info("Created CASE folder offer pending marked. workbook=" + workbookKey);
        }

        // メソッド: CASE/Base ブックの編集発生を記録する。
        // 引数: workbook - 対象ブック。
        // 戻り値: なし。
        // 副作用: セッション汚染状態を更新する。
        internal void HandleSheetChanged(Excel.Workbook workbook)
        {
            bool isBaseOrCaseWorkbook = IsBaseOrCaseWorkbookCore(workbook);
            bool isManagedClose = IsManagedCloseCore(workbook);
            bool isSuppressed = IsSuppressedCore(workbook);

            CaseWorkbookSheetChangeAction action = CaseWorkbookSheetChangePolicy.Decide(
                isBaseOrCaseWorkbook,
                isManagedClose,
                isSuppressed);

            if (action == CaseWorkbookSheetChangeAction.Ignore)
            {
                return;
            }

            if (action == CaseWorkbookSheetChangeAction.IgnoreBecauseTransientPaneSuppression)
            {
                _logger.Info("Case workbook sheet change ignored during transient suppression. workbook=" + GetWorkbookKey(workbook));
                return;
            }

            string workbookKey = GetWorkbookKey(workbook);
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _sessionDirtyWorkbookKeys.Add(workbookKey);
        }

        private bool IsBaseOrCaseWorkbook(Excel.Workbook workbook)
        {
            return _workbookRoleResolver.IsBaseWorkbook(workbook) || _workbookRoleResolver.IsCaseWorkbook(workbook);
        }

        private bool IsBaseOrCaseWorkbookCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.IsBaseOrCaseWorkbook != null
                ? _testHooks.IsBaseOrCaseWorkbook(workbook)
                : IsBaseOrCaseWorkbook(workbook);
        }

        private bool IsCaseWorkbookCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.IsCaseWorkbook != null
                ? _testHooks.IsCaseWorkbook(workbook)
                : _workbookRoleResolver.IsCaseWorkbook(workbook);
        }

        private void RegisterKnownCaseWorkbookCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.RegisterKnownCaseWorkbook != null)
            {
                _testHooks.RegisterKnownCaseWorkbook(workbook);
                return;
            }

            _workbookRoleResolver.RegisterKnownCaseWorkbook(workbook);
        }

        private void SyncNameRulesFromKernelToCaseCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.SyncNameRulesFromKernelToCase != null)
            {
                _testHooks.SyncNameRulesFromKernelToCase(workbook);
                return;
            }

            SyncNameRulesFromKernelToCase(workbook);
        }

        private bool IsManagedCloseCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.IsManagedClose != null
                ? _testHooks.IsManagedClose(workbook)
                : IsManagedClose(workbook);
        }

        private bool IsSuppressedCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.IsSuppressed != null
                ? _testHooks.IsSuppressed(workbook)
                : _transientPaneSuppressionService.IsSuppressed(workbook);
        }

        private DialogResult ShowClosePromptCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.ShowClosePrompt != null)
            {
                return _testHooks.ShowClosePrompt(workbook);
            }

            return _caseClosePromptService == null
                ? DialogResult.No
                : _caseClosePromptService.ShowClosePrompt(workbook);
        }

        private void ScheduleManagedSessionCloseCore(string workbookKey, string folderPath, bool saveChanges)
        {
            if (_testHooks != null && _testHooks.ScheduleManagedSessionClose != null)
            {
                _testHooks.ScheduleManagedSessionClose(workbookKey, folderPath, saveChanges);
                return;
            }

            ScheduleManagedSessionClose(workbookKey, folderPath, saveChanges);
        }

        private void SchedulePostCloseFollowUpCore(string workbookKey, string folderPath)
        {
            if (_testHooks != null && _testHooks.SchedulePostCloseFollowUp != null)
            {
                _testHooks.SchedulePostCloseFollowUp(workbookKey, folderPath);
                return;
            }

            SchedulePostCloseFollowUp(workbookKey, folderPath);
        }

        // メソッド: CASE HOME シート表示時に A 列が常に見えるようウィンドウ固定を再適用する。
        // 引数: workbook - 対象 CASE ブック。
        // 戻り値: なし。
        // 副作用: Excel Window の FreezePanes / ScrollRow / ScrollColumn を更新する。
        private void EnsureCaseHomeLeftColumnVisible(Excel.Workbook workbook)
        {
            if (!_workbookRoleResolver.IsCaseWorkbook(workbook))
            {
                return;
            }

            Excel.Worksheet activeWorksheet = workbook == null ? null : workbook.ActiveSheet as Excel.Worksheet;
            if (!IsCaseHomeSheet(activeWorksheet))
            {
                return;
            }

            Excel.Window targetWindow = ResolveCaseWindow(workbook);
            if (targetWindow == null)
            {
                return;
            }

            // 処理ブロック: 既存の固定状態を解除してから先頭 1 列固定を再設定し、A 列が隠れない状態を揃える。
            if (targetWindow.FreezePanes)
            {
                targetWindow.FreezePanes = false;
            }

            targetWindow.SplitRow = 0;
            targetWindow.SplitColumn = 1;
            targetWindow.FreezePanes = true;
            targetWindow.ScrollRow = 1;
            targetWindow.ScrollColumn = 1;
        }

        // メソッド: CASE HOME の画面制御対象となる Window を解決する。
        // 引数: workbook - 対象 CASE ブック。
        // 戻り値: 固定再適用先の Window。取得できない場合は null。
        // 副作用: なし。
        private Excel.Window ResolveCaseWindow(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();
            string activeWorkbookFullName = _excelInteropService.GetWorkbookFullName(activeWorkbook);
            string targetWorkbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
            if (string.Equals(activeWorkbookFullName, targetWorkbookFullName, StringComparison.OrdinalIgnoreCase))
            {
                Excel.Window activeWindow = _excelInteropService.GetActiveWindow();
                if (activeWindow != null)
                {
                    return activeWindow;
                }
            }

            return _excelInteropService.GetFirstVisibleWindow(workbook);
        }

        // メソッド: 指定シートが CASE HOME シートかを判定する。
        // 引数: worksheet - 判定対象シート。
        // 戻り値: CASE HOME シートなら true。
        // 副作用: なし。
        private static bool IsCaseHomeSheet(Excel.Worksheet worksheet)
        {
            return worksheet != null
                && string.Equals(worksheet.CodeName, CaseHomeSheetCodeName, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsManagedClose(Excel.Workbook workbook)
        {
            return _managedCloseState.IsManagedClose(GetWorkbookKey(workbook));
        }

        private void SchedulePostCloseFollowUp(string workbookKey, string folderPath)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _postCloseFollowUpScheduler?.Schedule(workbookKey, folderPath);
        }

        private void ScheduleManagedSessionClose(string workbookKey, string folderPath, bool saveChanges)
        {
            EnsureManagedCloseDispatcher().BeginInvoke((MethodInvoker)(() => ExecuteManagedSessionClose(workbookKey, folderPath, saveChanges)));
        }

        private void ExecuteManagedSessionClose(string workbookKey, string folderPath, bool saveChanges)
        {
            Excel.Workbook workbook = FindOpenWorkbook(workbookKey);
            if (workbook == null)
            {
                return;
            }

            try
            {
                using (BeginManagedCloseScope(workbook))
                {
                    if (saveChanges)
                    {
                        workbook.Save();
                    }
                    else
                    {
                        WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(workbook);
                    }

                    _sessionDirtyWorkbookKeys.Remove(workbookKey);
                    SchedulePostCloseFollowUp(workbookKey, folderPath);

                    bool previousDisplayAlerts = true;
                    bool hasDisplayAlertsSnapshot = false;
                    try
                    {
                        previousDisplayAlerts = _application.DisplayAlerts;
                        hasDisplayAlertsSnapshot = true;
                        _application.DisplayAlerts = false;
                        WorkbookCloseInteropHelper.CloseWithoutSave(workbook);
                    }
                    finally
                    {
                        if (hasDisplayAlertsSnapshot)
                        {
                            _application.DisplayAlerts = previousDisplayAlerts;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Case workbook managed session close failed.", ex);
                MessageBox.Show(
                    "保存または終了に失敗しました。もう一度お試しください。",
                    _caseClosePromptService == null ? "案件情報System" : _caseClosePromptService.GetCloseDialogTitle(workbook),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private Control EnsureManagedCloseDispatcher()
        {
            if (_managedCloseDispatcher != null && !_managedCloseDispatcher.IsDisposed)
            {
                return _managedCloseDispatcher;
            }

            _managedCloseDispatcher = new Control();
            IntPtr unusedHandle = _managedCloseDispatcher.Handle;
            return _managedCloseDispatcher;
        }

        private void SyncNameRulesFromKernelToCase(Excel.Workbook caseWorkbook)
        {
            if (_kernelNameRuleReader == null
                || !_kernelNameRuleReader.TryReadForCaseWorkbook(caseWorkbook, out string normalizedRuleA, out string normalizedRuleB))
            {
                return;
            }

            string currentRuleA = KernelNamingService.NormalizeNameRuleA(_excelInteropService.TryGetDocumentProperty(caseWorkbook, NameRuleAPropertyName));
            string currentRuleB = KernelNamingService.NormalizeNameRuleB(_excelInteropService.TryGetDocumentProperty(caseWorkbook, NameRuleBPropertyName));

            if (!string.Equals(currentRuleA, normalizedRuleA, StringComparison.OrdinalIgnoreCase))
            {
                _excelInteropService.SetDocumentProperty(caseWorkbook, NameRuleAPropertyName, normalizedRuleA);
            }

            if (!string.Equals(currentRuleB, normalizedRuleB, StringComparison.OrdinalIgnoreCase))
            {
                _excelInteropService.SetDocumentProperty(caseWorkbook, NameRuleBPropertyName, normalizedRuleB);
            }
        }

        private string ResolveContainingFolder(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.ResolveContainingFolder != null)
            {
                return _testHooks.ResolveContainingFolder(workbook) ?? string.Empty;
            }

            return _caseFolderOpenService == null
                ? string.Empty
                : _caseFolderOpenService.ResolveContainingFolder(workbook);
        }

        private bool DirectoryExistsSafe(string folderPath)
        {
            if (_testHooks != null && _testHooks.DirectoryExistsSafe != null)
            {
                return _testHooks.DirectoryExistsSafe(folderPath);
            }

            return _caseFolderOpenService != null
                && _caseFolderOpenService.DirectoryExistsSafe(folderPath);
        }

        private void PromptToOpenCreatedCaseFolderIfNeeded(string workbookKey, string folderPath)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            bool wasPending = _createdCaseFolderOfferPendingWorkbookKeys.Remove(workbookKey);
            if (!wasPending)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(folderPath) || !DirectoryExistsSafe(folderPath))
            {
                _logger.Info("Created CASE folder offer pending cleared without scheduling because folder was unavailable. workbook=" + workbookKey + ", folderPath=" + (folderPath ?? string.Empty));
                return;
            }

            _logger.Info("Created CASE folder offer pending will be prompted during before-close handling. workbook=" + workbookKey + ", folderPath=" + folderPath);
            TryPromptToOpenCreatedCaseFolder(folderPath);
        }

        private void TryPromptToOpenCreatedCaseFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return;
            }

            if (!DirectoryExistsSafe(folderPath))
            {
                _logger.Info("Created CASE folder offer prompt skipped because folder does not exist. folderPath=" + folderPath);
                return;
            }

            try
            {
                DialogResult answer = ShowCreatedCaseFolderOfferPromptCore(folderPath);
                if (answer != DialogResult.Yes)
                {
                    _logger.Info("Created CASE folder offer prompt dismissed without opening folder. folderPath=" + folderPath + ", answer=" + answer.ToString());
                    return;
                }

                OpenCreatedCaseFolderCore(folderPath);
            }
            catch (Exception ex)
            {
                _logger.Error("Created CASE folder offer prompt failed.", ex);
            }
        }

        private DialogResult ShowCreatedCaseFolderOfferPromptCore(string folderPath)
        {
            if (_testHooks != null && _testHooks.ShowCreatedCaseFolderOfferPrompt != null)
            {
                return _testHooks.ShowCreatedCaseFolderOfferPrompt(folderPath);
            }

            return _caseClosePromptService == null
                ? DialogResult.None
                : _caseClosePromptService.ShowCreatedCaseFolderOfferPrompt();
        }

        private void OpenCreatedCaseFolderCore(string folderPath)
        {
            if (_testHooks != null && _testHooks.OpenCreatedCaseFolder != null)
            {
                _testHooks.OpenCreatedCaseFolder(folderPath, "CaseWorkbookLifecycleService.PostCloseCreatedCaseFolderOffer");
                return;
            }

            _caseFolderOpenService?.OpenCreatedCaseFolder(folderPath);
        }

        private Excel.Workbook FindOpenWorkbook(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return null;
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (string.Equals(GetWorkbookKey(workbook), workbookKey, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }

            return null;
        }

        private string GetWorkbookKey(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.GetWorkbookKey != null)
            {
                return _testHooks.GetWorkbookKey(workbook) ?? string.Empty;
            }

            if (workbook == null)
            {
                return string.Empty;
            }

            string fullName = _excelInteropService.GetWorkbookFullName(workbook);
            return string.IsNullOrWhiteSpace(fullName)
                ? _excelInteropService.GetWorkbookName(workbook)
                : fullName;
        }

        private bool IsMacroEnabledWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                return workbook.FileFormat == Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
            }
            catch
            {
                return false;
            }
        }

        internal sealed class CaseWorkbookLifecycleServiceTestHooks
        {
            internal Func<Excel.Workbook, string> GetWorkbookKey { get; set; }

            internal Func<Excel.Workbook, bool> IsBaseOrCaseWorkbook { get; set; }

            internal Func<Excel.Workbook, bool> IsCaseWorkbook { get; set; }

            internal Action<Excel.Workbook> RegisterKnownCaseWorkbook { get; set; }

            internal Action<Excel.Workbook> SyncNameRulesFromKernelToCase { get; set; }

            internal Func<Excel.Workbook, bool> IsManagedClose { get; set; }

            internal Func<Excel.Workbook, bool> IsSuppressed { get; set; }

            internal Func<Excel.Workbook, string> ResolveContainingFolder { get; set; }

            internal Func<string, bool> DirectoryExistsSafe { get; set; }

            internal Func<Excel.Workbook, DialogResult> ShowClosePrompt { get; set; }

            internal Func<string, DialogResult> ShowCreatedCaseFolderOfferPrompt { get; set; }

            internal Action<string, string> OpenCreatedCaseFolder { get; set; }

            internal Action<string, string, bool> ScheduleManagedSessionClose { get; set; }

            internal Action<string, string> SchedulePostCloseFollowUp { get; set; }
        }

    }
}
