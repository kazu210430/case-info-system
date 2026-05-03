using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    // クラス: MasterList からテンプレート一覧を取得する。
    // 責務: SYSTEM_ROOT を手掛かりに Master ブックを開き、文書キーとテンプレート情報を一覧化する。
    /// <summary>
    /// Master ブックの雛形一覧を読み取り、文書ボタン用のテンプレート定義を返すサービス。
    /// SYSTEM_ROOT を基準に Master ブックを開き、一覧シートをキャッシュして利用する。
    /// </summary>
    internal sealed class MasterTemplateCatalogService : IMasterTemplateCatalogReader
    {
        private const string MasterSheetName = "雛形一覧";
        private const string MasterSheetCodeName = "shMasterList";
        private readonly ExcelInteropService _excelInteropService;
        private readonly MasterWorkbookReadAccessService _masterWorkbookReadAccessService;
        private readonly IMasterTemplateSheetReader _masterTemplateSheetReader;
        private readonly Logger _logger;
        private readonly Dictionary<string, IReadOnlyList<MasterTemplateRecord>> _cachedTemplatesByMasterPath =
            new Dictionary<string, IReadOnlyList<MasterTemplateRecord>>(StringComparer.OrdinalIgnoreCase);

        internal MasterTemplateCatalogService(
            ExcelInteropService excelInteropService,
            MasterWorkbookReadAccessService masterWorkbookReadAccessService,
            IMasterTemplateSheetReader masterTemplateSheetReader,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _masterWorkbookReadAccessService = masterWorkbookReadAccessService ?? throw new ArgumentNullException(nameof(masterWorkbookReadAccessService));
            _masterTemplateSheetReader = masterTemplateSheetReader ?? throw new ArgumentNullException(nameof(masterTemplateSheetReader));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        // メソッド: 読み取り済みテンプレート一覧キャッシュを破棄する。
        // 引数: なし。
        // 戻り値: なし。
        // 副作用: 内部キャッシュ状態を初期化する。
        internal void InvalidateCache()
        {
            _cachedTemplatesByMasterPath.Clear();
        }

        internal void InvalidateCache(Excel.Workbook workbook)
        {
            string resolvedMasterPath = _masterWorkbookReadAccessService.ResolveMasterPath(workbook, MasterWorkbookPathResolutionMode.MasterTemplateCatalog);
            if (resolvedMasterPath.Length == 0)
            {
                InvalidateCache();
                return;
            }

            _cachedTemplatesByMasterPath.Remove(resolvedMasterPath);
        }

        public bool TryGetTemplateByKey(Excel.Workbook caseWorkbook, string key, out MasterTemplateRecord record)
        {
            record = null;
            string normalizedKey = NormalizeDocButtonKey(key);
            if (normalizedKey.Length == 0)
            {
                return false;
            }

            IReadOnlyList<MasterTemplateRecord> templates = GetMasterTemplateList(caseWorkbook);
            for (int index = 0; index < templates.Count; index++)
            {
                MasterTemplateRecord current = templates[index];
                if (current != null && string.Equals(current.Key, normalizedKey, StringComparison.OrdinalIgnoreCase))
                {
                    record = current;
                    return true;
                }
            }

            return false;
        }

        private IReadOnlyList<MasterTemplateRecord> GetMasterTemplateList(Excel.Workbook caseWorkbook)
        {
            string resolvedMasterPath = _masterWorkbookReadAccessService.ResolveMasterPath(caseWorkbook, MasterWorkbookPathResolutionMode.MasterTemplateCatalog);
            if (resolvedMasterPath.Length == 0)
            {
                throw new InvalidOperationException("Master workbook path could not be resolved.");
            }

            if (_cachedTemplatesByMasterPath.TryGetValue(resolvedMasterPath, out IReadOnlyList<MasterTemplateRecord> cachedTemplates)
                && cachedTemplates != null)
            {
                return cachedTemplates;
            }

            MasterWorkbookReadAccessResult readAccess = _masterWorkbookReadAccessService.OpenReadOnly(
                resolvedMasterPath,
                MasterWorkbookOpenSearchMode.StrictFullPathOnly,
                path => new InvalidOperationException("Masterブックが見つかりません。 path=" + path));
            try
            {
                IReadOnlyList<MasterTemplateRecord> templates = ReadMasterTemplateList(readAccess.Workbook);
                _cachedTemplatesByMasterPath[resolvedMasterPath] = templates;
                return templates;
            }
            finally
            {
                readAccess.CloseIfOwned();
            }
        }

        // メソッド: MasterList の内容をテンプレート一覧へ変換する。
        // 引数: masterWorkbook - Master ブック。
        // 戻り値: テンプレート一覧。
        // 副作用: Worksheet / Range / Cell / Interior にアクセスする。
        private List<MasterTemplateRecord> ReadMasterTemplateList(Excel.Workbook masterWorkbook)
        {
            if (masterWorkbook == null)
            {
                throw new InvalidOperationException("Master workbook was null.");
            }

            Excel.Worksheet worksheet = GetMasterListWorksheet(masterWorkbook);
            if (worksheet == null)
            {
                throw new InvalidOperationException("雛形一覧シートが見つかりません。 book=" + _excelInteropService.GetWorkbookFullName(masterWorkbook));
            }

            MasterTemplateSheetData masterSheetData = _masterTemplateSheetReader.Read(worksheet);
            var result = new List<MasterTemplateRecord>();
            if (masterSheetData.LastRow < 3)
            {
                return result;
            }

            for (int index = 0; index < masterSheetData.Rows.Count; index++)
            {
                MasterTemplateSheetRowData row = masterSheetData.Rows[index];
                string key = row.Key;
                string templateFileName = row.TemplateFileName;
                string documentName = row.Caption;
                if (string.IsNullOrWhiteSpace(key) && string.IsNullOrWhiteSpace(templateFileName) && string.IsNullOrWhiteSpace(documentName))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(templateFileName))
                {
                    _logger.Warn("MasterTemplateCatalogService ignored incomplete row. row=" + row.RowIndex.ToString() + ", key=" + key + ", templateFileName=" + templateFileName);
                    continue;
                }

                result.Add(new MasterTemplateRecord
                {
                    Key = key,
                    TemplateFileName = templateFileName,
                    DocumentName = documentName,
                    BackColor = row.FillColor
                });
            }

            return result;
        }

        private Excel.Worksheet GetMasterListWorksheet(Excel.Workbook masterWorkbook)
        {
            Excel.Worksheet byCodeName = _excelInteropService.FindWorksheetByCodeName(masterWorkbook, MasterSheetCodeName);
            if (byCodeName != null)
            {
                return byCodeName;
            }

            try
            {
                return masterWorkbook.Worksheets[MasterSheetName] as Excel.Worksheet;
            }
            catch (Exception ex)
            {
                _logger.Error("MasterTemplateCatalogService.GetMasterListWorksheet failed.", ex);
                return null;
            }
        }

        private static string NormalizeDocButtonKey(string key)
        {
            string trimmed = (key ?? string.Empty).Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            return long.TryParse(trimmed, out long numericKey)
                ? numericKey.ToString("00")
                : trimmed;
        }
    }
}
