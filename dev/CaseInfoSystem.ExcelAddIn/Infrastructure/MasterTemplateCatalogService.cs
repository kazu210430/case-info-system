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
    internal sealed class MasterTemplateCatalogService : IMasterTemplateCatalogReader
    {
        private const string MasterSheetName = "雛形一覧";
        private const string MasterSheetCodeName = "shMasterList";
        private const string SystemRootPropertyName = "SYSTEM_ROOT";

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly IMasterTemplateSheetReader _masterTemplateSheetReader;
        private readonly Logger _logger;
        private readonly Dictionary<string, IReadOnlyList<MasterTemplateRecord>> _cachedTemplatesByMasterPath =
            new Dictionary<string, IReadOnlyList<MasterTemplateRecord>>(StringComparer.OrdinalIgnoreCase);

        internal MasterTemplateCatalogService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            IMasterTemplateSheetReader masterTemplateSheetReader,
            Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
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
            string resolvedMasterPath = ResolveMasterPath(workbook);
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
            string resolvedMasterPath = ResolveMasterPath(caseWorkbook);
            if (resolvedMasterPath.Length == 0)
            {
                throw new InvalidOperationException("Master workbook path could not be resolved.");
            }

            if (_cachedTemplatesByMasterPath.TryGetValue(resolvedMasterPath, out IReadOnlyList<MasterTemplateRecord> cachedTemplates)
                && cachedTemplates != null)
            {
                return cachedTemplates;
            }

            bool wasAlreadyOpen;
            Excel.Workbook masterWorkbook = OpenMasterReadOnly(resolvedMasterPath, out wasAlreadyOpen);
            try
            {
                IReadOnlyList<MasterTemplateRecord> templates = ReadMasterTemplateList(masterWorkbook);
                _cachedTemplatesByMasterPath[resolvedMasterPath] = templates;
                return templates;
            }
            finally
            {
                if (!wasAlreadyOpen && masterWorkbook != null)
                {
                    masterWorkbook.Close(false);
                }
            }
        }

        private Excel.Workbook OpenMasterReadOnly(string resolvedMasterPath, out bool wasAlreadyOpen)
        {
            wasAlreadyOpen = false;

            if (resolvedMasterPath.Length == 0)
            {
                throw new InvalidOperationException("Master workbook path could not be resolved.");
            }

            Excel.Workbook openWorkbook = FindOpenMasterWorkbook(resolvedMasterPath);
            if (openWorkbook != null)
            {
                wasAlreadyOpen = true;
                return openWorkbook;
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

        private Excel.Workbook FindOpenMasterWorkbook(string resolvedMasterPath)
        {
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (string.Equals(_pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookFullName(workbook)), resolvedMasterPath, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }

            return null;
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

        private static void HideWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            Excel.Windows windows = null;
            try
            {
                windows = workbook.Windows;
                int windowCount = windows == null ? 0 : windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = windows[index];
                        if (window != null)
                        {
                            window.Visible = false;
                        }
                    }
                    finally
                    {
                        if (window != null && Marshal.IsComObject(window))
                        {
                            ComObjectReleaseService.Release(window);
                        }
                    }
                }
            }
            catch
            {
                // 非表示化に失敗しても、Master 読み取り自体は継続する。
            }
            finally
            {
                if (windows != null && Marshal.IsComObject(windows))
                {
                    ComObjectReleaseService.Release(windows);
                }
            }
        }

        // メソッド: COM オブジェクトを解放する。
        // 引数: comObject - 解放対象。
        // 戻り値: なし。
        // 副作用: COM 参照を解放する。
        private static void ReleaseComObject(object comObject)
        {
            // Master catalog 読み取りで所有した COM 参照は完全解放の方針を維持する。
            ComObjectReleaseService.FinalRelease(comObject);
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
