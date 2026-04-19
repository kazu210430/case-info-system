using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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
    internal sealed class MasterTemplateCatalogService
    {
        private const string MasterSheetName = "雛形一覧";
        private const string MasterSheetCodeName = "shMasterList";
        private const int MasterListFirstDataRow = 3;
        private const string SystemRootPropertyName = "SYSTEM_ROOT";

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;
        private List<MasterTemplateRecord> _cachedTemplates;
        private bool _isCacheValid;

        internal MasterTemplateCatalogService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        // メソッド: 読み取り済みテンプレート一覧キャッシュを破棄する。
        // 引数: なし。
        // 戻り値: なし。
        // 副作用: 内部キャッシュ状態を初期化する。
        internal void InvalidateCache()
        {
            _isCacheValid = false;
            _cachedTemplates = null;
        }

        internal bool TryGetTemplateByKey(Excel.Workbook caseWorkbook, string key, out MasterTemplateRecord record)
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

        /// <summary>
        /// メソッド: MasterList に登録された文書テンプレート一覧を取得する。
        /// 引数: caseWorkbook - SYSTEM_ROOT を解決するための CASE ブック。
        /// 戻り値: 読み込み済みのテンプレート一覧。
        /// 副作用: 必要に応じて Master ブックを読み取り、キャッシュを更新する。
        /// </summary>
        internal IReadOnlyList<MasterTemplateRecord> GetAllTemplates(Excel.Workbook caseWorkbook)
        {
            return GetMasterTemplateList(caseWorkbook);
        }

        private IReadOnlyList<MasterTemplateRecord> GetMasterTemplateList(Excel.Workbook caseWorkbook)
        {
            if (_isCacheValid && _cachedTemplates != null)
            {
                return _cachedTemplates;
            }

            bool wasAlreadyOpen;
            string resolvedMasterPath;
            Excel.Workbook masterWorkbook = OpenMasterReadOnly(caseWorkbook, out wasAlreadyOpen, out resolvedMasterPath);
            try
            {
                _cachedTemplates = ReadMasterTemplateList(masterWorkbook);
                _isCacheValid = true;
                return _cachedTemplates;
            }
            finally
            {
                if (!wasAlreadyOpen && masterWorkbook != null)
                {
                    masterWorkbook.Close(false);
                }
            }
        }

        private Excel.Workbook OpenMasterReadOnly(Excel.Workbook caseWorkbook, out bool wasAlreadyOpen, out string resolvedMasterPath)
        {
            resolvedMasterPath = ResolveMasterPath(caseWorkbook);
            wasAlreadyOpen = false;

            if (resolvedMasterPath.Length == 0)
            {
                throw new InvalidOperationException("Master workbook path could not be resolved.");
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (string.Equals(_pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookFullName(workbook)), resolvedMasterPath, StringComparison.OrdinalIgnoreCase))
                {
                    wasAlreadyOpen = true;
                    HideWorkbookWindows(workbook);
                    return workbook;
                }
            }

            if (!_pathCompatibilityService.FileExistsSafe(resolvedMasterPath))
            {
                throw new InvalidOperationException("Masterブックが見つかりません。 path=" + resolvedMasterPath);
            }

            bool previousEnableEvents = _application.EnableEvents;
            try
            {
                _application.EnableEvents = false;
                Excel.Workbook workbook = _application.Workbooks.Open(resolvedMasterPath, ReadOnly: true, AddToMru: false, IgnoreReadOnlyRecommended: true, UpdateLinks: 0);
                HideWorkbookWindows(workbook);
                return workbook;
            }
            finally
            {
                _application.EnableEvents = previousEnableEvents;
            }
        }

        private string ResolveMasterPath(Excel.Workbook caseWorkbook)
        {
            string systemRoot = GetSystemRootFromBook(caseWorkbook);
            return systemRoot.Length == 0
                ? string.Empty
                : WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private string GetSystemRootFromBook(Excel.Workbook workbook)
        {
            string root = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName));
            if (root.Length > 0 && _pathCompatibilityService.FileExistsSafe(WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(root, _pathCompatibilityService)))
            {
                return root;
            }

            root = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(workbook));
            if (root.Length > 0 && _pathCompatibilityService.FileExistsSafe(WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(root, _pathCompatibilityService)))
            {
                return root;
            }

            string parentRoot = _pathCompatibilityService.NormalizePath(_pathCompatibilityService.GetParentFolderPath(_excelInteropService.GetWorkbookPath(workbook)));
            if (parentRoot.Length > 0 && _pathCompatibilityService.FileExistsSafe(WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(parentRoot, _pathCompatibilityService)))
            {
                return parentRoot;
            }

            return root;
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

            Excel.Range lastCell = worksheet.Cells[worksheet.Rows.Count, "A"] as Excel.Range;
            if (lastCell == null)
            {
                throw new InvalidOperationException("雛形一覧シートの最終行セルを取得できませんでした。");
            }

            Excel.Range lastUsedCell = lastCell.End[Excel.XlDirection.xlUp] as Excel.Range;
            if (lastUsedCell == null)
            {
                throw new InvalidOperationException("雛形一覧シートの使用済み最終行を取得できませんでした。");
            }

            int lastRow = lastUsedCell.Row;
            var result = new List<MasterTemplateRecord>();
            if (lastRow < MasterListFirstDataRow)
            {
                return result;
            }

            Excel.Range valuesRange = null;
            try
            {
                // 処理ブロック: A:C は一括取得し、色だけ個別セルから取得する。
                // テンプレート行数が多いほどセル単位読込の往復回数差が効くため、まず値の取得をまとめる。
                valuesRange = worksheet.Range["A" + MasterListFirstDataRow.ToString(), "C" + lastRow.ToString()];
                object[,] values = valuesRange.Value2 as object[,];
                if (values == null)
                {
                    return result;
                }

                int upperRow = values.GetUpperBound(0);
                for (int rowOffset = 1; rowOffset <= upperRow; rowOffset++)
                {
                    int rowIndex = MasterListFirstDataRow + rowOffset - 1;
                    string key = NormalizeDocButtonKey(Convert.ToString(values[rowOffset, 1]));
                    string templateFileName = Convert.ToString(values[rowOffset, 2]) ?? string.Empty;
                    string documentName = Convert.ToString(values[rowOffset, 3]) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(key) && string.IsNullOrWhiteSpace(templateFileName) && string.IsNullOrWhiteSpace(documentName))
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(templateFileName))
                    {
                        _logger.Warn("MasterTemplateCatalogService ignored incomplete row. row=" + rowIndex.ToString() + ", key=" + key + ", templateFileName=" + templateFileName);
                        continue;
                    }

                    long backColor = GetCellInteriorColor(worksheet, rowIndex, 4);
                    result.Add(new MasterTemplateRecord
                    {
                        Key = key,
                        TemplateFileName = templateFileName,
                        DocumentName = documentName,
                        BackColor = backColor
                    });
                }
            }
            finally
            {
                ReleaseComObject(valuesRange);
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

        // メソッド: 指定セルの塗りつぶし色を取得する。
        // 引数: worksheet - 対象シート, rowIndex - 行番号, columnIndex - 列番号。
        // 戻り値: OLE Color 値。
        // 副作用: Excel Cell / Interior にアクセスする。
        private static long GetCellInteriorColor(Excel.Worksheet worksheet, int rowIndex, int columnIndex)
        {
            Excel.Range cell = null;
            Excel.Interior interior = null;
            try
            {
                cell = worksheet.Cells[rowIndex, columnIndex] as Excel.Range;
                interior = cell == null ? null : cell.Interior;
                return Convert.ToInt64(interior == null ? 0 : interior.Color);
            }
            finally
            {
                ReleaseComObject(interior);
                ReleaseComObject(cell);
            }
        }

        private static void HideWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null)
                    {
                        window.Visible = false;
                    }
                }
            }
            catch
            {
                // 非表示化に失敗しても、Master 読み取り自体は継続する。
            }
        }

        // メソッド: COM オブジェクトを解放する。
        // 引数: comObject - 解放対象。
        // 戻り値: なし。
        // 副作用: COM 参照を解放する。
        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
                // 例外処理: COM 解放失敗は致命ではないため握りつぶす。
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
