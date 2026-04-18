using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // クラス: CASE/Base ブックの保存・終了まわりのライフサイクルを管理する。
    // 責務: 初期化、終了確認、管理対象クローズ、後追い終了処理を制御する。
    internal sealed class CaseWorkbookLifecycleService
    {
        private const string RolePropertyName = "ROLE";
        private const string CaseRoleName = "CASE";
        private const string BaseRoleName = "BASE";
        private const string CaseHomeSheetCodeName = "shHOME";
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string NameRuleAPropertyName = "NAME_RULE_A";
        private const string NameRuleBPropertyName = "NAME_RULE_B";
        private const string DefaultNameRuleA = "YYYY";
        private const string DefaultNameRuleB = "DOC";
        private const int ExcelBusyHResult = unchecked((int)0x800AC472);
        private const int PostCloseRetryCount = 20;
        private const int PostCloseRetryIntervalMs = 500;
        private const uint GuiCaretBlinking = 0x00000001;
        private static readonly XNamespace CustomPropertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly TransientPaneSuppressionService _transientPaneSuppressionService;
        private readonly Logger _logger;
        private readonly HashSet<string> _initializedWorkbookKeys;
        private readonly HashSet<string> _sessionDirtyWorkbookKeys;
        private readonly Dictionary<string, int> _managedCloseCounts;
        private readonly Queue<PostCloseFollowUpRequest> _pendingPostCloseQueue;
        private readonly CaseWorkbookLifecycleServiceTestHooks _testHooks;
        private Control _managedCloseDispatcher;
        private bool _postClosePosted;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GuiThreadInfo lpgui);

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            public uint cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public NativeRect rcCaret;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        internal CaseWorkbookLifecycleService(
            WorkbookRoleResolver workbookRoleResolver,
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            TransientPaneSuppressionService transientPaneSuppressionService,
            Logger logger)
        {
            _workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException(nameof(workbookRoleResolver));
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException(nameof(transientPaneSuppressionService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _initializedWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _sessionDirtyWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _managedCloseCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _pendingPostCloseQueue = new Queue<PostCloseFollowUpRequest>();
            _testHooks = null;
        }

        internal CaseWorkbookLifecycleService(Logger logger, CaseWorkbookLifecycleServiceTestHooks testHooks)
        {
            _workbookRoleResolver = null;
            _application = null;
            _excelInteropService = null;
            _pathCompatibilityService = null;
            _transientPaneSuppressionService = null;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _initializedWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _sessionDirtyWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _managedCloseCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _pendingPostCloseQueue = new Queue<PostCloseFollowUpRequest>();
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
            CaseWorkbookBeforeCloseAction action = CaseWorkbookBeforeClosePolicy.Decide(
                isBaseOrCaseWorkbook: IsBaseOrCaseWorkbookCore(workbook),
                isManagedClose: IsManagedCloseCore(workbook),
                isSessionDirty: _sessionDirtyWorkbookKeys.Contains(GetWorkbookKey(workbook)));

            if (action == CaseWorkbookBeforeCloseAction.Ignore)
            {
                return false;
            }

            if (action == CaseWorkbookBeforeCloseAction.SuppressPromptForManagedClose)
            {
                _logger.Info("Case workbook before-close prompt suppressed for managed close. workbook=" + GetWorkbookKey(workbook));
                return false;
            }

            try
            {
                string workbookKey = GetWorkbookKey(workbook);
                string folderPath = ResolveContainingFolder(workbook);

                if (action == CaseWorkbookBeforeCloseAction.PromptForDirtySession)
                {
                    cancel = true;

                    DialogResult answer = ShowClosePromptCore(workbook);

                    if (answer == DialogResult.Cancel)
                    {
                        return true;
                    }

                    ScheduleManagedSessionCloseCore(workbookKey, folderPath, answer == DialogResult.Yes);
                    return true;
                }

                _logger.Info(
                    "Case workbook before-close follow-up is handled by VSTO. workbook="
                    + workbookKey
                    + ", macroEnabled="
                    + IsMacroEnabledWorkbook(workbook).ToString());

                SchedulePostCloseFollowUpCore(workbookKey, folderPath);
                _logger.Info("Case workbook post-close follow-up posted. workbook=" + workbookKey);
                return false;
            }
            catch (Exception ex)
            {
                // 例外処理: 終了イベントで例外を再送出すると Excel のクローズ動作が不安定になるため、ログ化して返す。
                _logger.Error("Case workbook before-close handling failed.", ex);
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
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return NoOpDisposable.Instance;
            }

            if (_managedCloseCounts.ContainsKey(workbookKey))
            {
                _managedCloseCounts[workbookKey] = _managedCloseCounts[workbookKey] + 1;
            }
            else
            {
                _managedCloseCounts.Add(workbookKey, 1);
            }

            return new ManagedCloseScope(this, workbookKey);
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
            _managedCloseCounts.Remove(workbookKey);
            _workbookRoleResolver.RemoveKnownWorkbook(workbook);
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

            return MessageBox.Show(
                "保存しますか？",
                BuildCloseDialogTitle(workbook),
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);
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
            string workbookKey = GetWorkbookKey(workbook);
            return workbookKey.Length > 0
                && _managedCloseCounts.TryGetValue(workbookKey, out int count)
                && count > 0;
        }

        private void ReleaseManagedClose(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey) || !_managedCloseCounts.TryGetValue(workbookKey, out int count))
            {
                return;
            }

            if (count <= 1)
            {
                _managedCloseCounts.Remove(workbookKey);
                return;
            }

            _managedCloseCounts[workbookKey] = count - 1;
        }

        private void SchedulePostCloseFollowUp(string workbookKey, string folderPath)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _pendingPostCloseQueue.Enqueue(new PostCloseFollowUpRequest(workbookKey, folderPath, PostCloseRetryCount));
            if (_postClosePosted)
            {
                return;
            }

            _postClosePosted = true;
            EnsureManagedCloseDispatcher().BeginInvoke((MethodInvoker)ExecutePendingPostCloseQueue);
        }

        // メソッド: 終了後フォローアップの待機キューを順に処理する。
        // 引数: なし。
        // 戻り値: なし。
        // 副作用: Excel 終了判定、再試行予約、ログ出力を行う。
        private void ExecutePendingPostCloseQueue()
        {
            _postClosePosted = false;

            while (_pendingPostCloseQueue.Count > 0)
            {
                PostCloseFollowUpRequest request = _pendingPostCloseQueue.Dequeue();
                if (request == null)
                {
                    continue;
                }

                try
                {
                    if (IsWorkbookStillOpen(request.WorkbookKey))
                    {
                        _logger.Info("Case workbook post-close follow-up skipped because workbook is still open. workbook=" + request.WorkbookKey);
                        continue;
                    }

                    QuitExcelIfNoVisibleWorkbook();
                }
                catch (COMException ex) when (ex.ErrorCode == ExcelBusyHResult && request.AttemptsRemaining > 0)
                {
                    _logger.Info(
                        "Case workbook post-close follow-up will retry because Excel is busy. workbook="
                        + request.WorkbookKey
                        + ", attemptsRemaining="
                        + request.AttemptsRemaining.ToString());
                    _pendingPostCloseQueue.Enqueue(request.NextAttempt());
                    SchedulePendingPostCloseRetry();
                    return;
                }
                catch (Exception ex)
                {
                    // 例外処理: 後追いフォローアップは補助処理のため、失敗しても本体のクローズ結果を優先して継続する。
                    _logger.Error("Case workbook post-close follow-up failed.", ex);
                }
            }
        }

        private void SchedulePendingPostCloseRetry()
        {
            if (_postClosePosted)
            {
                return;
            }

            _postClosePosted = true;
            Timer retryTimer = new Timer();
            retryTimer.Interval = PostCloseRetryIntervalMs;
            retryTimer.Tick += (sender, args) =>
            {
                retryTimer.Stop();
                retryTimer.Dispose();
                ExecutePendingPostCloseQueue();
            };
            retryTimer.Start();
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
                    bool previousDisplayAlerts = _application.DisplayAlerts;
                    try
                    {
                        _application.DisplayAlerts = false;

                        if (saveChanges)
                        {
                            workbook.Save();
                        }
                        else
                        {
                            workbook.Saved = true;
                        }

                        _sessionDirtyWorkbookKeys.Remove(workbookKey);
                        SchedulePostCloseFollowUp(workbookKey, folderPath);
                        workbook.Close(SaveChanges: false);
                    }
                    finally
                    {
                        _application.DisplayAlerts = previousDisplayAlerts;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Case workbook managed session close failed.", ex);
                MessageBox.Show(
                    "保存または終了に失敗しました。もう一度お試しください。",
                    BuildCloseDialogTitle(workbook),
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

        private void QuitExcelIfNoVisibleWorkbook()
        {
            bool hasVisibleWorkbook = false;
            foreach (Excel.Workbook openWorkbook in _application.Workbooks)
            {
                if (openWorkbook == null)
                {
                    continue;
                }

                try
                {
                    if (openWorkbook.Windows.Count > 0 && openWorkbook.Windows.Cast<Excel.Window>().Any(window => window.Visible))
                    {
                        hasVisibleWorkbook = true;
                        break;
                    }
                }
                catch
                {
                    // Closing workbook may already be tearing down. Ignore and keep scanning.
                }
            }

            _logger.Info("Case post-close visible workbook check. hasVisibleWorkbook=" + hasVisibleWorkbook.ToString());
            if (hasVisibleWorkbook)
            {
                return;
            }

            bool previousDisplayAlerts = _application.DisplayAlerts;
            try
            {
                _application.DisplayAlerts = false;
                _logger.Info("Case post-close quitting Excel because no visible workbook remains.");
                _application.Quit();
            }
            finally
            {
                _application.DisplayAlerts = previousDisplayAlerts;
            }
        }

        private void SyncNameRulesFromKernelToCase(Excel.Workbook caseWorkbook)
        {
            string kernelPath = ResolveKernelWorkbookPath(caseWorkbook);
            if (string.IsNullOrWhiteSpace(kernelPath) || !_pathCompatibilityService.FileExistsSafe(kernelPath))
            {
                return;
            }

            if (!TryGetKernelNameRules(kernelPath, out string kernelRuleA, out string kernelRuleB))
            {
                return;
            }

            string normalizedRuleA = KernelNamingService.NormalizeNameRuleA(kernelRuleA);
            string normalizedRuleB = KernelNamingService.NormalizeNameRuleB(kernelRuleB);

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

        private string ResolveKernelWorkbookPath(Excel.Workbook caseWorkbook)
        {
            string systemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(caseWorkbook, SystemRootPropertyName));
            if (systemRoot.Length == 0)
            {
                return string.Empty;
            }

            return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private bool TryGetKernelNameRules(string kernelPath, out string ruleA, out string ruleB)
        {
            ruleA = DefaultNameRuleA;
            ruleB = DefaultNameRuleB;

            Excel.Workbook openKernelWorkbook = _excelInteropService.FindOpenWorkbook(kernelPath);
            if (openKernelWorkbook != null)
            {
                ruleA = _excelInteropService.TryGetDocumentProperty(openKernelWorkbook, NameRuleAPropertyName);
                ruleB = _excelInteropService.TryGetDocumentProperty(openKernelWorkbook, NameRuleBPropertyName);
                return true;
            }

            return TryReadKernelNameRulesFromPackage(kernelPath, out ruleA, out ruleB);
        }

        private bool TryReadKernelNameRulesFromPackage(string kernelPath, out string ruleA, out string ruleB)
        {
            ruleA = DefaultNameRuleA;
            ruleB = DefaultNameRuleB;

            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(kernelPath))
                {
                    ZipArchiveEntry customXmlEntry = archive.GetEntry("docProps/custom.xml");
                    if (customXmlEntry == null)
                    {
                        return false;
                    }

                    using (Stream stream = customXmlEntry.Open())
                    {
                        XDocument document = XDocument.Load(stream);
                        foreach (XElement propertyElement in document.Root == null ? Array.Empty<XElement>() : document.Root.Elements(CustomPropertiesNamespace + "property"))
                        {
                            XAttribute nameAttribute = propertyElement.Attribute("name");
                            if (nameAttribute == null)
                            {
                                continue;
                            }

                            XElement valueElement = propertyElement.Elements().FirstOrDefault();
                            if (valueElement == null)
                            {
                                continue;
                            }

                            string propertyValue = valueElement.Value ?? string.Empty;
                            if (string.Equals(nameAttribute.Value, NameRuleAPropertyName, StringComparison.OrdinalIgnoreCase))
                            {
                                ruleA = propertyValue;
                            }
                            else if (string.Equals(nameAttribute.Value, NameRuleBPropertyName, StringComparison.OrdinalIgnoreCase))
                            {
                                ruleB = propertyValue;
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel package name-rule read failed. path=" + kernelPath, ex);
                return false;
            }
        }

        private string ResolveContainingFolder(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.ResolveContainingFolder != null)
            {
                return _testHooks.ResolveContainingFolder(workbook) ?? string.Empty;
            }

            string folderPath = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(workbook));
            if (folderPath.Length == 0 || !_pathCompatibilityService.DirectoryExistsSafe(folderPath))
            {
                return string.Empty;
            }

            return folderPath;
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

        private string BuildCloseDialogTitle(Excel.Workbook workbook)
        {
            string role = (_excelInteropService.TryGetDocumentProperty(workbook, RolePropertyName) ?? string.Empty).Trim().ToUpperInvariant();
            if (role == CaseRoleName)
            {
                return "案件情報System (CASE)";
            }

            if (role == BaseRoleName)
            {
                return "案件情報System (BASE)";
            }

            return "案件情報System";
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

        private bool IsWorkbookStillOpen(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return false;
            }

            foreach (Excel.Workbook openWorkbook in _application.Workbooks)
            {
                if (openWorkbook == null)
                {
                    continue;
                }

                if (string.Equals(GetWorkbookKey(openWorkbook), workbookKey, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
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

        private sealed class ManagedCloseScope : IDisposable
        {
            private readonly CaseWorkbookLifecycleService _owner;
            private readonly string _workbookKey;
            private bool _disposed;

            internal ManagedCloseScope(CaseWorkbookLifecycleService owner, string workbookKey)
            {
                _owner = owner;
                _workbookKey = workbookKey;
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _owner.ReleaseManagedClose(_workbookKey);
            }
        }

        private sealed class PostCloseFollowUpRequest
        {
            internal PostCloseFollowUpRequest(string workbookKey, string folderPath, int attemptsRemaining = 3)
            {
                WorkbookKey = workbookKey ?? string.Empty;
                FolderPath = folderPath ?? string.Empty;
                AttemptsRemaining = attemptsRemaining;
            }

            internal string WorkbookKey { get; }

            internal string FolderPath { get; }

            internal int AttemptsRemaining { get; }

            internal PostCloseFollowUpRequest NextAttempt()
            {
                return new PostCloseFollowUpRequest(WorkbookKey, FolderPath, AttemptsRemaining - 1);
            }
        }

        private sealed class NoOpDisposable : IDisposable
        {
            internal static readonly NoOpDisposable Instance = new NoOpDisposable();

            public void Dispose()
            {
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

            internal Func<Excel.Workbook, DialogResult> ShowClosePrompt { get; set; }

            internal Action<string, string, bool> ScheduleManagedSessionClose { get; set; }

            internal Action<string, string> SchedulePostCloseFollowUp { get; set; }
        }

    }
}
