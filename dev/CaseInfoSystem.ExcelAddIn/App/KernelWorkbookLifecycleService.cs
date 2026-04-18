using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// Kernel ブックの保存前同期と終了処理を扱うサービス。
    /// DocProp 更新、保存確認ダイアログ、managed close をまとめて制御する。
    /// </summary>
    internal sealed class KernelWorkbookLifecycleService
    {
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string WordTemplateDirectoryPropertyName = "WORD_TEMPLATE_DIR";
        private const string TemplateFolderName = "雛形";

        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;
        private readonly Dictionary<string, int> _managedCloseCounts;
        private Control _managedCloseDispatcher;
        private int _beforeSaveDocPropSynchronizationSuppressionCount;

        /// <summary>
        /// コンストラクタ: KernelWorkbookLifecycleService を初期化する。
        /// 引数: Workbook 制御に必要な各サービスと Excel Application を受け取る。
        /// </summary>
        internal KernelWorkbookLifecycleService(
            KernelWorkbookService kernelWorkbookService,
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger)
        {
            _kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException(nameof(kernelWorkbookService));
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _managedCloseCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// メソッド: Kernel 保存前に DocProp を同期する。
        /// 引数: workbook は対象ブック、saveAsUi は名前を付けて保存 UI 利用有無、cancel は保存キャンセル可否。
        /// </summary>
        internal void HandleWorkbookBeforeSave(Excel.Workbook workbook, bool saveAsUi, ref bool cancel)
        {
            if (cancel || !_kernelWorkbookService.IsKernelWorkbook(workbook))
            {
                return;
            }

            if (_beforeSaveDocPropSynchronizationSuppressionCount > 0)
            {
                _logger.Info("Kernel workbook before-save docprop synchronization suppressed. workbook=" + GetWorkbookKey(workbook));
                return;
            }

            if (IsTransientReadOnlyKernelWorkbook(workbook))
            {
                _logger.Info("Kernel workbook before-save synchronization skipped for transient read-only workbook. workbook=" + GetWorkbookKey(workbook));
                return;
            }

            try
            {
                string workbookPath = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(workbook));
                if (workbookPath.Length == 0)
                {
                    return;
                }

                _excelInteropService.SetDocumentProperty(workbook, SystemRootPropertyName, workbookPath);
                _excelInteropService.SetDocumentProperty(
                    workbook,
                    WordTemplateDirectoryPropertyName,
                    _pathCompatibilityService.CombinePath(workbookPath, TemplateFolderName));

                _logger.Info(
                    "Kernel workbook before-save docprops synchronized. path="
                    + workbookPath
                    + ", saveAsUi="
                    + saveAsUi.ToString());
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel workbook before-save synchronization failed.", ex);
            }
        }

        /// <summary>
        /// メソッド: Kernel 保存前の DocProp 自動同期を一時抑止するスコープを開始する。
        /// 引数: reason - ログ用理由。
        /// 戻り値: スコープ解除用 IDisposable。
        /// 副作用: 保存前同期抑止カウンタを更新する。
        /// </summary>
        internal IDisposable SuppressBeforeSaveDocPropSynchronization(string reason)
        {
            _beforeSaveDocPropSynchronizationSuppressionCount++;
            _logger.Info(
                "Kernel before-save docprop synchronization suppression entered. reason="
                + (reason ?? string.Empty)
                + ", count="
                + _beforeSaveDocPropSynchronizationSuppressionCount.ToString());
            return new DelegateDisposable(() =>
            {
                if (_beforeSaveDocPropSynchronizationSuppressionCount > 0)
                {
                    _beforeSaveDocPropSynchronizationSuppressionCount--;
                }

                _logger.Info(
                    "Kernel before-save docprop synchronization suppression exited. reason="
                    + (reason ?? string.Empty)
                    + ", count="
                    + _beforeSaveDocPropSynchronizationSuppressionCount.ToString());
            });
        }
        /// <summary>
        /// メソッド: Kernel ブック終了時の保存確認と managed close への切替を行う。
        /// 引数: workbook は終了対象、cancel は終了キャンセル可否。
        /// 戻り値: このハンドラで終了処理を引き受けた場合 true。
        /// </summary>
        internal bool HandleWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (!_kernelWorkbookService.IsKernelWorkbook(workbook))
            {
                return false;
            }
            if (IsManagedClose(workbook))
            {
                _logger.Info("Kernel workbook before-close prompt suppressed for managed close. workbook=" + GetWorkbookKey(workbook));
                return false;
            }

            if (IsTransientReadOnlyKernelWorkbook(workbook))
            {
                try
                {
                    WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(workbook);
                    _logger.Info("Kernel workbook before-close prompt suppressed for transient read-only workbook. workbook=" + GetWorkbookKey(workbook));
                }
                catch (Exception ex)
                {
                    _logger.Error("Kernel workbook transient read-only close preparation failed.", ex);
                }

                return false;
            }

            if (!RequiresSave(workbook))
            {
                return false;
            }

            try
            {
                cancel = true;
                DialogResult answer = MessageBox.Show(
                    "保存しますか？",
                    BuildCloseDialogTitle(workbook),
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);

                if (answer == DialogResult.Cancel)
                {
                    return true;
                }

                ScheduleManagedClose(GetWorkbookKey(workbook), answer == DialogResult.Yes);
                return true;
            }
            catch (Exception ex)
            {
                // 例外処理: 終了イベントで例外を再送出すると Excel のクローズ動作が不安定になるため、ログ化して通常動作へ戻す。
                _logger.Error("Kernel workbook before-close handling failed.", ex);
                return false;
            }
        }

        /// <summary>
        /// メソッド: HOME 終了要求を managed close として予約する。
        /// </summary>

        /// <summary>
        /// メソッド: Kernel HOME 終了時の close 要求を managed close に統一する。
        /// 引数: workbook - 終了対象の Kernel Workbook。
        /// 戻り値: close 継続または予約できた場合 true。利用者がキャンセルした場合 false。
        /// 副作用: 必要に応じて保存確認ダイアログ表示と managed close 予約を行う。
        /// </summary>
        internal bool RequestManagedCloseFromHomeExit(Excel.Workbook workbook)
        {
            if (!_kernelWorkbookService.IsKernelWorkbook(workbook))
            {
                return false;
            }

            try
            {
                bool requiresSave = RequiresSave(workbook);
                DialogResult answer = DialogResult.No;

                if (requiresSave)
                {
                    answer = MessageBox.Show(
                        "保存しますか？",
                        BuildCloseDialogTitle(workbook),
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1);

                    if (answer == DialogResult.Cancel)
                    {
                        _logger.Info("Kernel HOME exit managed close canceled by user. workbook=" + GetWorkbookKey(workbook));
                        return false;
                    }
                }

                bool saveChanges = requiresSave && answer == DialogResult.Yes;
                ScheduleManagedClose(GetWorkbookKey(workbook), saveChanges);
                _logger.Info(
                    "Kernel HOME exit managed close scheduled. workbook="
                    + GetWorkbookKey(workbook)
                    + ", requiresSave="
                    + requiresSave.ToString()
                    + ", saveChanges="
                    + saveChanges.ToString());
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel HOME exit managed close scheduling failed.", ex);
                return false;
            }
        }
        private void ScheduleManagedClose(string workbookKey, bool saveChanges)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            EnsureManagedCloseDispatcher().BeginInvoke((MethodInvoker)(() => ExecuteManagedClose(workbookKey, saveChanges)));
        }

        /// <summary>
        /// メソッド: 一時的に開いた読み取り専用の Kernel ブックかを判定する。
        /// 引数: workbook - 判定対象の Kernel Workbook。
        /// 戻り値: 一時参照用途の読み取り専用ブックなら true。
        /// 副作用: なし。
        /// </summary>
        private bool IsTransientReadOnlyKernelWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                if (!workbook.ReadOnly)
                {
                    return false;
                }

                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null && window.Visible)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("Transient read-only kernel workbook detection failed.", ex);
                return false;
            }
        }

        /// <summary>
        /// メソッド: managed close として保存と close を実行する。
        /// </summary>
        private void ExecuteManagedClose(string workbookKey, bool saveChanges)
        {
            Excel.Workbook workbook = _excelInteropService.FindOpenWorkbook(workbookKey);
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
                            WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(workbook);
                        }

                        workbook.Close(SaveChanges: false);
                        QuitExcelIfKernelWasLastWorkbook(workbook);
                    }
                    finally
                    {
                        _application.DisplayAlerts = previousDisplayAlerts;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel workbook managed close failed.", ex);
                MessageBox.Show(
                    "保存または終了に失敗しました。もう一度お試しください。",
                    BuildCloseDialogTitle(workbook),
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// メソッド: Kernel が最後のブックなら Excel を終了する。
        /// </summary>
        private void QuitExcelIfKernelWasLastWorkbook(Excel.Workbook closingWorkbook)
        {
            bool hasOtherVisibleWorkbook = HasOtherVisibleWorkbook(closingWorkbook);
            bool hasOtherWorkbook = HasOtherWorkbook(closingWorkbook);
            _logger.Info(
                "Kernel managed close post-check. workbook="
                + GetWorkbookKey(closingWorkbook)
                + ", hasOtherVisibleWorkbook="
                + hasOtherVisibleWorkbook.ToString()
                + ", hasOtherWorkbook="
                + hasOtherWorkbook.ToString());

            if (hasOtherVisibleWorkbook || hasOtherWorkbook)
            {
                return;
            }

            _logger.Info("Kernel managed close will quit Excel because no other workbook remains.");
            _application.Quit();
        }

        /// <summary>
        /// メソッド: managed close 中であることを示すスコープを開始する。
        /// </summary>
        private IDisposable BeginManagedCloseScope(Excel.Workbook workbook)
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

        /// <summary>
        /// メソッド: 現在 managed close 中かを判定する。
        /// </summary>
        private bool IsManagedClose(Excel.Workbook workbook)
        {
            string workbookKey = GetWorkbookKey(workbook);
            return workbookKey.Length > 0
                && _managedCloseCounts.TryGetValue(workbookKey, out int count)
                && count > 0;
        }

        /// <summary>
        /// メソッド: managed close の参照カウントを解放する。
        /// </summary>
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

        /// <summary>
        /// メソッド: 指定ブック以外に表示中 Workbook があるかを判定する。
        /// </summary>
        private bool HasOtherVisibleWorkbook(Excel.Workbook workbookToIgnore)
        {
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                {
                    continue;
                }

                if (workbook.Windows.Count > 0)
                {
                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (window != null && window.Visible)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// メソッド: 指定ブック以外の open Workbook が存在するかを判定する。
        /// </summary>
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

        /// <summary>
        /// メソッド: managed close 用のディスパッチ Control を用意する。
        /// </summary>
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

        /// <summary>
        /// メソッド: ブック識別用のキーを取得する。
        /// </summary>
        private string GetWorkbookKey(Excel.Workbook workbook)
        {
            return _excelInteropService.GetWorkbookFullName(workbook);
        }

        /// <summary>
        /// メソッド: close ダイアログのタイトル文字列を組み立てる。
        /// </summary>
        private static string BuildCloseDialogTitle(Excel.Workbook workbook)
        {
            string workbookName = workbook == null ? string.Empty : workbook.Name;
            if (string.IsNullOrWhiteSpace(workbookName))
            {
                return "案件情報System";
            }

            return workbookName;
        }

        /// <summary>
        /// メソッド: ブックに未保存変更があるかを判定する。
        /// </summary>
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
                // 例外処理: Saved 判定に失敗した場合は安全側で保存確認ありとして扱う。
                return true;
            }
        }

        /// <summary>
        /// managed close スコープ終了時に参照カウントを解放する IDisposable 実装。
        /// </summary>
        private sealed class ManagedCloseScope : IDisposable
        {
            private readonly KernelWorkbookLifecycleService _owner;
            private readonly string _workbookKey;
            private bool _disposed;

            internal ManagedCloseScope(KernelWorkbookLifecycleService owner, string workbookKey)
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

        /// <summary>
        /// 何もしない IDisposable 実装。
        /// </summary>
        private sealed class NoOpDisposable : IDisposable
        {
            internal static readonly NoOpDisposable Instance = new NoOpDisposable();

            public void Dispose()
            {
            }
        }

        /// <summary>
        /// クラス: 任意処理を Dispose 時に実行する。
        /// 責務: 簡易な suppression scope を提供する。
        /// </summary>
        private sealed class DelegateDisposable : IDisposable
        {
            private readonly Action _disposeAction;
            private bool _disposed;

            /// <summary>
            /// メソッド: DelegateDisposable を初期化する。
            /// 引数: disposeAction - 解放時処理。
            /// 戻り値: なし。
            /// 副作用: なし。
            /// </summary>
            internal DelegateDisposable(Action disposeAction)
            {
                _disposeAction = disposeAction ?? throw new ArgumentNullException(nameof(disposeAction));
            }

            /// <summary>
            /// メソッド: 解放処理を一度だけ実行する。
            /// 引数: なし。
            /// 戻り値: なし。
            /// 副作用: 抑止カウンタ更新処理を実行する。
            /// </summary>
            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _disposeAction();
            }
        }
    }
}




