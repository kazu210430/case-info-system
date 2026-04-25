using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// クラス: Word 文書作成フロー全体を調停する。
    /// 責務: テンプレ解決、案件値収集、差込、保存、Word 表示までを VSTO 主導で実行する。
    /// </summary>
    internal sealed class DocumentCreateService
    {
        private const string DocumentNameOverrideEnabledPropertyName = "TASKPANE_DOC_NAME_OVERRIDE_ENABLED";
        private const string DocumentNameOverridePropertyName = "TASKPANE_DOC_NAME_OVERRIDE";

        private readonly ExcelInteropService _excelInteropService;
        private readonly CaseContextFactory _caseContextFactory;
        private readonly DocumentOutputService _documentOutputService;
        private readonly MergeDataBuilder _mergeDataBuilder;
        private readonly DocumentMergeService _documentMergeService;
        private readonly DocumentSaveService _documentSaveService;
        private readonly WordInteropService _wordInteropService;
        private readonly DocumentPresentationWaitService _documentPresentationWaitService;
        private readonly Logger _logger;

        /// <summary>
        /// メソッド: サービスを初期化する。
        /// 引数: excelInteropService - Excel 操作サービス, caseContextFactory - CASE 文脈生成サービス,
        /// documentTemplateResolver - テンプレ解決サービス, documentOutputService - 保存先解決サービス,
        /// mergeDataBuilder - 差込データ生成サービス, documentMergeService - 文書差込サービス,
        /// wordInteropService - Word 操作サービス, logger - ロガー。
        /// 戻り値: なし。
        /// 副作用: なし。
        /// </summary>
        internal DocumentCreateService(
            ExcelInteropService excelInteropService,
            CaseContextFactory caseContextFactory,
            DocumentOutputService documentOutputService,
            MergeDataBuilder mergeDataBuilder,
            DocumentMergeService documentMergeService,
            DocumentSaveService documentSaveService,
            WordInteropService wordInteropService,
            DocumentPresentationWaitService documentPresentationWaitService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _caseContextFactory = caseContextFactory ?? throw new ArgumentNullException(nameof(caseContextFactory));
            _documentOutputService = documentOutputService ?? throw new ArgumentNullException(nameof(documentOutputService));
            _mergeDataBuilder = mergeDataBuilder ?? throw new ArgumentNullException(nameof(mergeDataBuilder));
            _documentMergeService = documentMergeService ?? throw new ArgumentNullException(nameof(documentMergeService));
            _documentSaveService = documentSaveService ?? throw new ArgumentNullException(nameof(documentSaveService));
            _wordInteropService = wordInteropService ?? throw new ArgumentNullException(nameof(wordInteropService));
            _documentPresentationWaitService = documentPresentationWaitService ?? throw new ArgumentNullException(nameof(documentPresentationWaitService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// メソッド: 解決済みテンプレートを使って VSTO 文書作成を実行する。
        /// 引数: workbook - 対象 CASE ブック, templateSpec - 解決済みテンプレート。
        /// 戻り値: なし。
        /// 副作用: Word 文書を生成・保存・表示する。
        /// </summary>
        internal void Execute(Excel.Workbook workbook, DocumentTemplateSpec templateSpec)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (templateSpec == null)
            {
                throw new ArgumentNullException(nameof(templateSpec));
            }

            CaseContext caseContext = _caseContextFactory.CreateForDocumentCreate(workbook);
            Execute(workbook, templateSpec, caseContext);
        }

        /// <summary>
        /// メソッド: 解決済み CASE 文脈を再利用して VSTO 文書作成を実行する。
        /// 引数: workbook - 対象 CASE ブック, templateSpec - 解決済みテンプレート, caseContext - 押下時に解決済みの CASE 文脈。
        /// 戻り値: なし。
        /// 副作用: Word 文書を生成・保存・表示する。
        /// </summary>
        internal void Execute(Excel.Workbook workbook, DocumentTemplateSpec templateSpec, CaseContext caseContext)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (templateSpec == null)
            {
                throw new ArgumentNullException(nameof(templateSpec));
            }

            if (caseContext == null)
            {
                caseContext = _caseContextFactory.CreateForDocumentCreate(workbook);
            }

            if (caseContext == null || caseContext.CaseValues == null || caseContext.CaseValues.Count == 0)
            {
                throw new InvalidOperationException("案件データを管理定義に従って解決できませんでした。");
            }

            var totalStopwatch = Stopwatch.StartNew();
            var phaseStopwatch = Stopwatch.StartNew();
            string documentName = ResolveDocumentName(workbook, templateSpec);
            _logger.Debug("DocumentCreateService.Prepare", "DocumentNameResolved elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed));
            phaseStopwatch.Restart();
            string outputPath = _documentOutputService.BuildDocumentOutputPath(workbook, documentName, caseContext.CustomerName);
            _logger.Debug("DocumentCreateService.Prepare", "OutputPathResolved elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed) + " output=" + (outputPath ?? string.Empty));
            phaseStopwatch.Restart();
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new InvalidOperationException("出力先パスを作成できませんでした。");
            }

            IReadOnlyDictionary<string, string> mergeData = _mergeDataBuilder.BuildMergeData(caseContext);
            _logger.Debug("DocumentCreateService.Prepare", "MergeDataBuilt elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed) + " mergeFieldCount=" + mergeData.Count.ToString());
            ExecuteWordCreate(workbook, templateSpec, outputPath, mergeData, documentName);
        }

        /// <summary>
        /// メソッド: 文書名 override を考慮して最終文書名を解決する。
        /// 引数: workbook - 対象 CASE ブック, templateSpec - テンプレ情報。
        /// 戻り値: 最終文書名。
        /// 副作用: DocProp を読み取る。
        /// </summary>
        private string ResolveDocumentName(Excel.Workbook workbook, DocumentTemplateSpec templateSpec)
        {
            string defaultDocumentName = templateSpec == null ? string.Empty : (templateSpec.DocumentName ?? string.Empty);
            string overrideEnabled = _excelInteropService.TryGetDocumentProperty(workbook, DocumentNameOverrideEnabledPropertyName);
            bool isOverrideEnabled = string.Equals(overrideEnabled, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(overrideEnabled, "true", StringComparison.OrdinalIgnoreCase);
            if (!isOverrideEnabled)
            {
                return defaultDocumentName;
            }

            string overriddenName = (_excelInteropService.TryGetDocumentProperty(workbook, DocumentNameOverridePropertyName) ?? string.Empty).Trim();
            return overriddenName.Length == 0 ? defaultDocumentName : overriddenName;
        }

        /// <summary>
        /// メソッド: Word 文書作成の実処理を行う。
        /// 引数: workbook - 対象 CASE ブック, templateSpec - テンプレ情報, outputPath - 保存先パス, mergeData - 差込データ, documentName - 文書名。
        /// 戻り値: なし。
        /// 副作用: Word を起動し文書を作成・保存する。
        /// </summary>
        private void ExecuteWordCreate(
            Excel.Workbook workbook,
            DocumentTemplateSpec templateSpec,
            string outputPath,
            IReadOnlyDictionary<string, string> mergeData,
            string documentName)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            object wordApplication = null;
            object wordDocument = null;
            WordInteropService.WordPerformanceState wordPerformanceState = null;
            bool createdNewWord = false;
            string savedPath = string.Empty;
            string stage = "Initialize";
            ExcelUiState excelUiState = null;
            XlWindowState previousWindowState = XlWindowState.xlNormal;
            bool restoreExcelWindowPresentation = true;
            DocumentPresentationWaitService.WaitSession waitSession = null;
            var totalStopwatch = Stopwatch.StartNew();
            var phaseStopwatch = Stopwatch.StartNew();

            try
            {
                _logger.Debug("ExecuteCreateDocument", "Start template=" + (templateSpec == null ? string.Empty : templateSpec.TemplatePath) + " output=" + (outputPath ?? string.Empty));
                previousWindowState = Globals.ThisAddIn.Application.WindowState;
                excelUiState = ExcelUiState.Capture(Globals.ThisAddIn.Application);
                excelUiState.ApplyForDocumentCreate(Globals.ThisAddIn.Application, false);
                waitSession = _documentPresentationWaitService.ShowWaiting(totalStopwatch);

                stage = "AcquireWordApplication";
                SetStatusBar("文書作成：Word準備中...");
                waitSession?.UpdateStage(DocumentPresentationWaitService.LaunchingWordStageTitle);
                wordApplication = _wordInteropService.AcquireWordApplication(out createdNewWord);
                if (wordApplication == null)
                {
                    throw new InvalidOperationException("Word を起動または取得できませんでした。");
                }

                wordApplication = _wordInteropService.EnsureWordApplication(ref wordApplication);
                wordPerformanceState = _wordInteropService.BeginPerformanceScope(wordApplication, true, createdNewWord);
                _logger.Debug("ExecuteCreateDocument", "WordReady createdNew=" + createdNewWord.ToString() + " elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed));
                phaseStopwatch.Restart();

                stage = "CreateDocumentFromTemplate";
                SetStatusBar("文書作成：テンプレから作成中...");
                waitSession?.UpdateStage(DocumentPresentationWaitService.LoadingTemplateStageTitle);
                wordDocument = _wordInteropService.CreateDocumentFromTemplate(wordApplication, templateSpec.TemplatePath);
                if (wordDocument == null)
                {
                    throw new InvalidOperationException("テンプレートから Word 文書を作成できませんでした。");
                }
                _logger.Debug("ExecuteCreateDocument", "DocumentCreated elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed));
                phaseStopwatch.Restart();

                stage = "ApplyMergeData";
                SetStatusBar("文書作成：差し込み中...");
                waitSession?.UpdateStage(DocumentPresentationWaitService.ApplyingMergeDataStageTitle);
                _documentMergeService.ApplyMergeData(wordDocument, mergeData);
                _logger.Debug("ExecuteCreateDocument", "MergeApplied elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed));
                phaseStopwatch.Restart();

                stage = "RemoveContentControls";
                SetStatusBar("文書作成：仕上げ中...");
                _documentMergeService.RemoveContentControlsKeepText(wordDocument);
                _logger.Debug("ExecuteCreateDocument", "ControlsRemoved elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed));
                phaseStopwatch.Restart();

                stage = "SaveDocument";
                SetStatusBar("文書作成：保存中...");
                waitSession?.UpdateStage(DocumentPresentationWaitService.SavingDocumentStageTitle);
                DocumentSaveResult saveResult = _documentSaveService.SaveDocument(wordApplication, wordDocument, outputPath);
                if (saveResult == null || saveResult.ActiveDocument == null)
                {
                    throw new InvalidOperationException("保存後の Word 文書を再取得できませんでした。");
                }
                wordDocument = saveResult.ActiveDocument;
                savedPath = saveResult.FinalPath;
                _logger.Debug("ExecuteCreateDocument", "Saved path=" + (savedPath ?? string.Empty) + " elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed));
                phaseStopwatch.Restart();

                stage = "ShowDocument";
                SetStatusBar("文書作成：完了（Word表示）...");
                waitSession?.UpdateStage(DocumentPresentationWaitService.ShowingScreenStageTitle);
                _wordInteropService.ShowDocument(wordApplication, wordDocument);
                if (wordPerformanceState != null)
                {
                    // 表示後は Visible を元に戻さない。戻すと Word が再び非表示になる。
                    wordPerformanceState.HasVisible = false;
                }
                waitSession?.Close();
                restoreExcelWindowPresentation = false;
                _logger.Debug("ExecuteCreateDocument", "ShowDocument elapsed=" + FormatElapsedSeconds(phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds(totalStopwatch.Elapsed));
                _logger.Info(
                    "DocumentCreateService completed. documentName="
                    + (documentName ?? string.Empty)
                    + ", template="
                    + (templateSpec == null ? string.Empty : (templateSpec.TemplatePath ?? string.Empty))
                    + ", output="
                    + savedPath);
            }
            catch (Exception ex)
            {
                Exception innerException = ex.InnerException;
                _logger.Warn(
                    "DocumentCreateService.ExecuteWordCreate context."
                    + " stage=" + (stage ?? string.Empty)
                    + ", createdNewWord=" + createdNewWord.ToString()
                    + ", template=" + (templateSpec == null ? string.Empty : (templateSpec.TemplatePath ?? string.Empty))
                    + ", output=" + (outputPath ?? string.Empty)
                    + ", savedPath=" + (savedPath ?? string.Empty)
                    + ", exceptionType=" + ex.GetType().FullName
                    + ", message=" + ex.Message
                    + ", hresult=0x" + ex.HResult.ToString("X8")
                    + ", innerType=" + (innerException == null ? "(none)" : innerException.GetType().FullName)
                    + ", innerMessage=" + (innerException == null ? "(none)" : innerException.Message)
                    + ", innerHresult=" + (innerException == null ? "(none)" : "0x" + innerException.HResult.ToString("X8")));
                _logger.Error("DocumentCreateService.ExecuteWordCreate failed.", ex);
                _wordInteropService.RestorePerformanceState(wordApplication, wordPerformanceState);
                _wordInteropService.CloseDocumentNoSave(ref wordDocument);
                if (createdNewWord)
                {
                    _wordInteropService.QuitApplicationNoSave(ref wordApplication);
                }

                throw new InvalidOperationException(
                    "Word出力に失敗しました。" + Environment.NewLine
                    + "Template=" + (templateSpec == null ? string.Empty : (templateSpec.TemplatePath ?? string.Empty)) + Environment.NewLine
                    + "Output=" + (outputPath ?? string.Empty),
                    ex);
            }
            finally
            {
                _wordInteropService.RestorePerformanceState(wordApplication, wordPerformanceState);
                if (excelUiState != null)
                {
                    excelUiState.Restore(Globals.ThisAddIn.Application, restoreExcelWindowPresentation);
                }

                if (restoreExcelWindowPresentation)
                {
                    try
                    {
                        Globals.ThisAddIn.Application.WindowState = previousWindowState;
                    }
                    catch
                    {
                        // 例外処理: Excel ウィンドウ状態復元失敗は致命ではないため握りつぶす。
                    }
                }


                waitSession?.Dispose();
                ClearStatusBar();
            }
        }
        /// <summary>
        /// メソッド: Excel のステータスバーへ進捗を表示する。
        /// 引数: messageText - 表示文字列。
        /// 戻り値: なし。
        /// 副作用: ステータスバーを書き換える。
        /// </summary>
        private static void SetStatusBar(string messageText)
        {
            try
            {
                Globals.ThisAddIn.Application.StatusBar = messageText ?? string.Empty;
            }
            catch
            {
                // 例外処理: ステータスバー更新失敗は致命ではないため握りつぶす。
            }
        }

        /// <summary>
        /// メソッド: Excel のステータスバー表示を解除する。
        /// 引数: なし。
        /// 戻り値: なし。
        /// 副作用: ステータスバー制御を Excel 既定へ戻す。
        /// </summary>
        private static void ClearStatusBar()
        {
            try
            {
                Globals.ThisAddIn.Application.StatusBar = false;
            }
            catch
            {
                // 例外処理: ステータスバー解除失敗は致命ではないため握りつぶす。
            }
        }

        /// <summary>
        /// メソッド: 経過時間を見やすい文字列へ整形する。
        /// 引数: elapsed - 経過時間。
        /// 戻り値: 秒表記文字列。
        /// 副作用: なし。
        /// </summary>
        private static string FormatElapsedSeconds(TimeSpan elapsed)
        {
            return elapsed.TotalSeconds.ToString("0.000");
        }

        /// <summary>
        /// クラス: VBA の TExcelUIState 互換で Excel UI 状態を保持する。
        /// 責務: 実行前の ScreenUpdating / EnableEvents / DisplayAlerts / Calculation / Visible を保存復元する。
        /// </summary>
        private sealed class ExcelUiState
        {
            private ExcelUiState()
            {
            }

            internal bool ScreenUpdating { get; private set; }

            internal bool EnableEvents { get; private set; }

            internal bool DisplayAlerts { get; private set; }

            internal XlCalculation Calculation { get; private set; }

            internal bool Visible { get; private set; }

            internal static ExcelUiState Capture(Excel.Application application)
            {
                if (application == null)
                {
                    return null;
                }

                return new ExcelUiState
                {
                    ScreenUpdating = application.ScreenUpdating,
                    EnableEvents = application.EnableEvents,
                    DisplayAlerts = application.DisplayAlerts,
                    Calculation = application.Calculation,
                    Visible = application.Visible
                };
            }

            internal void ApplyForDocumentCreate(Excel.Application application, bool hideExcel)
            {
                if (application == null)
                {
                    return;
                }

                try
                {
                    application.ScreenUpdating = false;
                    application.EnableEvents = false;
                    application.DisplayAlerts = false;
                    application.Calculation = XlCalculation.xlCalculationManual;
                    if (hideExcel)
                    {
                        application.Visible = false;
                    }
                }
                catch
                {
                    // 例外処理: UI 設定変更失敗は致命ではないため握りつぶす。
                }
            }

            internal void Restore(Excel.Application application, bool restoreWindowPresentation)
            {
                if (application == null)
                {
                    return;
                }

                try
                {
                    application.ScreenUpdating = ScreenUpdating;
                    application.EnableEvents = EnableEvents;
                    application.DisplayAlerts = DisplayAlerts;
                    application.Calculation = Calculation;
                    if (restoreWindowPresentation)
                    {
                        application.Visible = Visible;
                    }
                }
                catch
                {
                    // 例外処理: UI 状態復元失敗は致命ではないため握りつぶす。
                }
            }
        }
    }
}


