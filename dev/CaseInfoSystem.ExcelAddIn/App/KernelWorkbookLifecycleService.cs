using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
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
        private static readonly string[] ProtectedManagementSheetNames = new string[]
        {
            "CaseList_FieldInventory",
            "CaseList_Headers",
            "CaseList_Mapping",
            "UserData_BaseMapping"
        };

        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly CaseListFieldDefinitionRepository _caseListFieldDefinitionRepository;
        private readonly CaseListHeaderRepository _caseListHeaderRepository;
        private readonly CaseListMappingRepository _caseListMappingRepository;
        private readonly Logger _logger;
        private readonly Dictionary<string, int> _managedCloseCounts;
        private Control _managedCloseDispatcher;
        private int _beforeSaveDocPropSynchronizationSuppressionCount;
        private Action<string, Excel.Workbook, bool> _homeManagedCloseStarted;
        private Action<string, Excel.Workbook, bool> _homeManagedCloseSucceeded;
        private Action<string, Excel.Workbook, bool, Exception> _homeManagedCloseFailed;

        /// <summary>
        /// コンストラクタ: KernelWorkbookLifecycleService を初期化する。
        /// 引数: Workbook 制御に必要な各サービスと Excel Application を受け取る。
        /// </summary>
        internal KernelWorkbookLifecycleService(
            KernelWorkbookService kernelWorkbookService,
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            CaseListFieldDefinitionRepository caseListFieldDefinitionRepository,
            CaseListHeaderRepository caseListHeaderRepository,
            CaseListMappingRepository caseListMappingRepository,
            Logger logger)
        {
            _kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException(nameof(kernelWorkbookService));
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException(nameof(caseListFieldDefinitionRepository));
            _caseListHeaderRepository = caseListHeaderRepository ?? throw new ArgumentNullException(nameof(caseListHeaderRepository));
            _caseListMappingRepository = caseListMappingRepository ?? throw new ArgumentNullException(nameof(caseListMappingRepository));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _managedCloseCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        }

        internal void RegisterHomeManagedCloseCallbacks(
            Action<string, Excel.Workbook, bool> onStarted,
            Action<string, Excel.Workbook, bool> onSucceeded,
            Action<string, Excel.Workbook, bool, Exception> onFailed)
        {
            _homeManagedCloseStarted = onStarted;
            _homeManagedCloseSucceeded = onSucceeded;
            _homeManagedCloseFailed = onFailed;
        }

        /// <summary>
        /// メソッド: Kernel ブック利用開始時に管理シート保護と軽量な定義整合確認を行う。
        /// 引数: workbook - 対象の Kernel ブック。
        /// 戻り値: なし。
        /// 副作用: hidden 管理シート保護を再適用し、整合不良はログへ警告する。
        /// </summary>
        internal void HandleWorkbookOpenedOrActivated(Excel.Workbook workbook)
        {
            if (!_kernelWorkbookService.IsKernelWorkbook(workbook))
            {
                return;
            }

            try
            {
                EnsureProtectedManagementSheets(workbook);

                string validationMessage = ValidateCaseListManagedDefinitions(workbook);
                if (!string.IsNullOrWhiteSpace(validationMessage))
                {
                    _logger.Warn("Kernel workbook case-list managed-definition validation failed. " + validationMessage);
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel workbook initialization guards failed.", ex);
            }
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
                Exception missingWorkbookException = new InvalidOperationException("Managed close target workbook was not found.");
                _logger.Error(
                    "Kernel managed close failed because workbook was not found. workbook="
                    + (workbookKey ?? string.Empty)
                    + ", saveChanges="
                    + saveChanges.ToString(),
                    missingWorkbookException);
                _homeManagedCloseFailed?.Invoke(workbookKey, null, saveChanges, missingWorkbookException);
                return;
            }

            string currentManagedCloseMethod = "BeginManagedCloseScope";
            string failedManagedCloseMethod = null;
            try
            {
                _logger.Info(
                    "Kernel managed close started. workbook="
                    + GetWorkbookKey(workbook)
                    + ", saveChanges="
                    + saveChanges.ToString());
                _homeManagedCloseStarted?.Invoke(workbookKey, workbook, saveChanges);
                using (BeginManagedCloseScope(workbook))
                {
                    if (saveChanges)
                    {
                        currentManagedCloseMethod = "Workbook.Save";
                        _logger.Info(
                            "Kernel managed close calling Workbook.Save. workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                        workbook.Save();
                        _logger.Info(
                            "Kernel managed close completed Workbook.Save. workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                    }
                    else
                    {
                        currentManagedCloseMethod = "WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose";
                        WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(workbook);
                    }

                    ExcelApplicationStateScope closeScope = new ExcelApplicationStateScope(_application);
                    try
                    {
                        currentManagedCloseMethod = "ExcelApplicationStateScope.SetDisplayAlerts(false)";
                        _logger.Info(
                            "Kernel managed close calling ExcelApplicationStateScope.SetDisplayAlerts(false). workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                        closeScope.SetDisplayAlerts(false);
                        _logger.Info(
                            "Kernel managed close completed ExcelApplicationStateScope.SetDisplayAlerts(false). workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                        currentManagedCloseMethod = "Workbook.Close(false, Type.Missing, Type.Missing)";
                        _logger.Info(
                            "Kernel managed close calling Workbook.Close(false, Type.Missing, Type.Missing). workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                        WorkbookCloseInteropHelper.CloseWithoutSave(workbook);
                        _logger.Info(
                            "Kernel managed close completed Workbook.Close(false, Type.Missing, Type.Missing). workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                    }
                    catch
                    {
                        failedManagedCloseMethod = currentManagedCloseMethod;
                        throw;
                    }
                    finally
                    {
                        _logger.Info(
                            "Kernel managed close restoring DisplayAlerts scope. workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                        try
                        {
                            closeScope.Dispose();
                        }
                        catch
                        {
                            if (string.IsNullOrWhiteSpace(failedManagedCloseMethod))
                            {
                                failedManagedCloseMethod = "ExcelApplicationStateScope.Dispose";
                            }

                            throw;
                        }
                        _logger.Info(
                            "Kernel managed close restored DisplayAlerts scope. workbook="
                            + GetWorkbookKey(workbook)
                            + ", saveChanges="
                            + saveChanges.ToString());
                    }

                    currentManagedCloseMethod = nameof(QuitExcelIfKernelWasLastWorkbook);
                    QuitExcelIfKernelWasLastWorkbook(workbook);
                }

                _logger.Info(
                    "Kernel managed close succeeded. workbook="
                    + GetWorkbookKey(workbook)
                    + ", saveChanges="
                    + saveChanges.ToString());
                _homeManagedCloseSucceeded?.Invoke(workbookKey, workbook, saveChanges);
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Kernel managed close failed. workbook="
                    + (workbookKey ?? string.Empty)
                    + ", saveChanges="
                    + saveChanges.ToString()
                    + ", failedMethod="
                    + (failedManagedCloseMethod ?? currentManagedCloseMethod)
                    + ", exceptionType="
                    + (ex.GetType().FullName ?? string.Empty)
                    + ", exceptionMessage="
                    + (ex.Message ?? string.Empty),
                    ex);
                _homeManagedCloseFailed?.Invoke(workbookKey, workbook, saveChanges, ex);
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
            ExcelApplicationStateScope quitScope = new ExcelApplicationStateScope(_application);
            try
            {
                quitScope.SetDisplayAlerts(false);
                _application.Quit();
            }
            finally
            {
                quitScope.Dispose();
            }
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

        private void EnsureProtectedManagementSheets(Excel.Workbook workbook)
        {
            foreach (string sheetName in ProtectedManagementSheetNames)
            {
                Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByName(workbook, sheetName);
                if (worksheet == null)
                {
                    _logger.Warn("Kernel workbook management sheet was not found. sheet=" + sheetName);
                    continue;
                }

                if (worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios)
                {
                    continue;
                }

                worksheet.Protect(
                    Password: string.Empty,
                    UserInterfaceOnly: true,
                    DrawingObjects: Type.Missing,
                    Contents: Type.Missing,
                    Scenarios: Type.Missing);
                worksheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;
                _logger.Info("Kernel workbook management sheet protected for runtime safety. sheet=" + sheetName);
            }
        }

        private string ValidateCaseListManagedDefinitions(Excel.Workbook workbook)
        {
            Excel.Worksheet caseListWorksheet = _excelInteropService.FindCaseListWorksheet(workbook);
            if (caseListWorksheet == null)
            {
                return "Kernelブックにシート「案件一覧」が見つかりません。";
            }

            IReadOnlyDictionary<string, CaseListFieldDefinition> fieldDefinitions = _caseListFieldDefinitionRepository.LoadDefinitions(workbook);
            IReadOnlyList<CaseListHeaderDefinition> headerDefinitions = _caseListHeaderRepository.LoadDefinitions(workbook);
            IReadOnlyList<CaseListMappingDefinition> enabledMappings = _caseListMappingRepository.LoadEnabledDefinitions(workbook);

            if (fieldDefinitions == null || fieldDefinitions.Count == 0)
            {
                return "Kernelブックの管理シート CaseList_FieldInventory を読み取れません。";
            }

            if (headerDefinitions == null || headerDefinitions.Count == 0)
            {
                return "Kernelブックの管理シート CaseList_Headers を読み取れません。";
            }

            if (enabledMappings == null || enabledMappings.Count == 0)
            {
                return "Kernelブックの管理シート CaseList_Mapping を読み取れません。";
            }

            IReadOnlyDictionary<string, int> managedHeaderMap = BuildManagedHeaderMap(headerDefinitions);
            IReadOnlyDictionary<string, int> actualHeaderMap = ReadActualCaseListHeaderMap(caseListWorksheet);
            foreach (KeyValuePair<string, int> pair in managedHeaderMap)
            {
                int actualColumn;
                if (!actualHeaderMap.TryGetValue(pair.Key, out actualColumn))
                {
                    return "案件一覧シートに管理定義ヘッダが存在しません。 header=" + pair.Key;
                }

                if (actualColumn != pair.Value)
                {
                    return "案件一覧シートの列配置が管理定義と一致しません。 header=" + pair.Key + ", managedColumn=" + pair.Value + ", actualColumn=" + actualColumn;
                }
            }

            foreach (CaseListMappingDefinition mapping in enabledMappings)
            {
                if (mapping == null)
                {
                    continue;
                }

                string sourceFieldKey = (mapping.SourceFieldKey ?? string.Empty).Trim();
                string targetHeaderName = (mapping.TargetHeaderName ?? string.Empty).Trim();
                if (!fieldDefinitions.ContainsKey(sourceFieldKey))
                {
                    return "CaseList_Mapping に FieldInventory 未定義の項目があります。 sourceFieldKey=" + sourceFieldKey;
                }

                if (!managedHeaderMap.ContainsKey(targetHeaderName))
                {
                    return "CaseList_Mapping に Headers 未定義のヘッダがあります。 targetHeaderName=" + targetHeaderName;
                }
            }

            return string.Empty;
        }

        private static IReadOnlyDictionary<string, int> BuildManagedHeaderMap(IReadOnlyList<CaseListHeaderDefinition> headerDefinitions)
        {
            Dictionary<string, int> result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (headerDefinitions == null)
            {
                return result;
            }

            foreach (CaseListHeaderDefinition definition in headerDefinitions)
            {
                string headerName = ((definition == null ? string.Empty : definition.HeaderName) ?? string.Empty).Trim();
                int columnIndex = ConvertColumnAddressToIndex(definition == null ? string.Empty : definition.CellAddress);
                if (headerName.Length == 0 || columnIndex <= 0 || result.ContainsKey(headerName))
                {
                    continue;
                }

                result.Add(headerName, columnIndex);
            }

            return result;
        }

        private static IReadOnlyDictionary<string, int> ReadActualCaseListHeaderMap(Excel.Worksheet caseListWorksheet)
        {
            Dictionary<string, int> result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Excel.Range range = null;
            try
            {
                int lastColumn = ((dynamic)caseListWorksheet.Cells[2, caseListWorksheet.Columns.Count]).End[Excel.XlDirection.xlToLeft].Column;
                if (lastColumn < 1)
                {
                    return result;
                }

                range = caseListWorksheet.Range[(dynamic)caseListWorksheet.Cells[2, 1], (dynamic)caseListWorksheet.Cells[2, lastColumn]];
                object[,] values = range.Value2 as object[,];
                if (values == null)
                {
                    return result;
                }

                int upperBound = values.GetUpperBound(1);
                for (int i = 1; i <= upperBound; i++)
                {
                    string headerName = (Convert.ToString(values[1, i]) ?? string.Empty).Trim();
                    if (headerName.Length != 0 && !result.ContainsKey(headerName))
                    {
                        result.Add(headerName, i);
                    }
                }

                return result;
            }
            finally
            {
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(range);
            }
        }

        private static int ConvertColumnAddressToIndex(string cellAddress)
        {
            if (string.IsNullOrWhiteSpace(cellAddress))
            {
                return 0;
            }

            int result = 0;
            string normalized = cellAddress.Trim().ToUpperInvariant();
            foreach (char c in normalized)
            {
                if (c < 'A' || c > 'Z')
                {
                    break;
                }

                result = result * 26 + (c - 'A' + 1);
            }

            return result;
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




