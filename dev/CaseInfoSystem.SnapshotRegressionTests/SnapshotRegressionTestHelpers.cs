using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.SnapshotRegressionTests
{
    internal sealed class SnapshotBuilderScenario : IDisposable
    {
        internal sealed class InputRow
        {
            internal string Key { get; set; } = string.Empty;
            internal string TemplateFileName { get; set; } = string.Empty;
            internal string Caption { get; set; } = string.Empty;
            internal string TabName { get; set; } = string.Empty;
            internal long FillColor { get; set; }
            internal long TabBackColor { get; set; }
        }

        internal Excel.Application Application { get; }

        internal Excel.Workbook CaseWorkbook { get; }

        internal Excel.Workbook MasterWorkbook { get; }

        internal TaskPaneSnapshotBuilderService Builder { get; }

        internal Array Values { get; }

        internal IReadOnlyDictionary<int, long> FillColors { get; }

        internal IReadOnlyDictionary<int, long> TabBackColors { get; }

        private readonly ExcelInteropService _excelInteropService;

        private SnapshotBuilderScenario(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            Excel.Workbook caseWorkbook,
            Excel.Workbook masterWorkbook,
            TaskPaneSnapshotBuilderService builder,
            Array values,
            IReadOnlyDictionary<int, long> fillColors,
            IReadOnlyDictionary<int, long> tabBackColors)
        {
            Application = application;
            _excelInteropService = excelInteropService;
            CaseWorkbook = caseWorkbook;
            MasterWorkbook = masterWorkbook;
            Builder = builder;
            Values = values;
            FillColors = fillColors;
            TabBackColors = tabBackColors;
        }

        internal static SnapshotBuilderScenario Create(
            IReadOnlyList<InputRow> rows,
            int masterVersion,
            bool caseListRegistered,
            Excel.Application application = null,
            string systemRoot = @"C:\SnapshotRegression\SystemRoot",
            string caseWorkbookPath = @"C:\SnapshotRegression\Cases\案件情報_山田.xlsx",
            string masterWorkbookPath = @"C:\SnapshotRegression\SystemRoot\案件情報System_Kernel.xlsx")
        {
            application ??= new Excel.Application();
            var logger = new Logger(_ => { });
            var pathCompatibilityService = new PathCompatibilityService();
            var excelInteropService = new ExcelInteropService(application, logger, pathCompatibilityService);
            var masterWorkbookReadAccessService = new MasterWorkbookReadAccessService(application, excelInteropService, pathCompatibilityService);
            var masterTemplateSheetReader = new MasterTemplateSheetReaderAdapter();
            var builder = new TaskPaneSnapshotBuilderService(excelInteropService, pathCompatibilityService, masterWorkbookReadAccessService, masterTemplateSheetReader, logger);

            var caseWorkbook = new Excel.Workbook
            {
                FullName = caseWorkbookPath,
                Name = "案件情報_山田.xlsx",
                Path = @"C:\SnapshotRegression\Cases",
                CustomDocumentProperties = new DocumentProperties()
            };
            excelInteropService.SetDocumentProperty(caseWorkbook, "SYSTEM_ROOT", systemRoot);
            excelInteropService.SetDocumentProperty(caseWorkbook, "CASELIST_REGISTERED", caseListRegistered ? "1" : "0");
            application.Workbooks.Add(caseWorkbook);

            var masterWorkbook = new Excel.Workbook
            {
                FullName = masterWorkbookPath,
                Name = "案件情報System_Kernel.xlsx",
                Path = systemRoot,
                CustomDocumentProperties = new DocumentProperties()
            };
            excelInteropService.SetDocumentProperty(masterWorkbook, "TASKPANE_MASTER_VERSION", masterVersion.ToString(CultureInfo.InvariantCulture));

            var masterWorksheet = new Excel.Worksheet
            {
                CodeName = "shMasterList",
                Name = "雛形一覧"
            };

            Array values = Array.CreateInstance(typeof(object), new[] { rows.Count, 5 }, new[] { 1, 1 });
            var fillColors = new Dictionary<int, long>();
            var tabBackColors = new Dictionary<int, long>();
            for (int index = 0; index < rows.Count; index++)
            {
                InputRow row = rows[index] ?? new InputRow();
                int worksheetRow = index + 3;
                values.SetValue(row.Key, index + 1, 1);
                values.SetValue(row.TemplateFileName, index + 1, 2);
                values.SetValue(row.Caption, index + 1, 3);
                values.SetValue(string.Empty, index + 1, 4);
                values.SetValue(row.TabName, index + 1, 5);

                masterWorksheet.Cells[worksheetRow, "A"].Value2 = row.Key;
                masterWorksheet.Cells[worksheetRow, "B"].Value2 = row.TemplateFileName;
                masterWorksheet.Cells[worksheetRow, "C"].Value2 = row.Caption;
                masterWorksheet.Cells[worksheetRow, "D"].Value2 = string.Empty;
                masterWorksheet.Cells[worksheetRow, "E"].Value2 = row.TabName;
                masterWorksheet.Cells[worksheetRow, "D"].Interior.Color = row.FillColor;
                masterWorksheet.Cells[worksheetRow, "F"].Interior.Color = row.TabBackColor;

                fillColors[worksheetRow] = row.FillColor;
                tabBackColors[worksheetRow] = row.TabBackColor;
            }

            masterWorkbook.Worksheets.Add(masterWorksheet);
            application.Workbooks.Add(masterWorkbook);

            return new SnapshotBuilderScenario(
                application,
                excelInteropService,
                caseWorkbook,
                masterWorkbook,
                builder,
                values,
                fillColors,
                tabBackColors);
        }

        internal string LoadCaseCacheSnapshot()
        {
            string rawCount = _excelInteropService.TryGetDocumentProperty(CaseWorkbook, "TASKPANE_SNAPSHOT_CACHE_COUNT");
            if (!int.TryParse(rawCount, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) || count <= 0)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            for (int index = 1; index <= count; index++)
            {
                builder.Append(_excelInteropService.TryGetDocumentProperty(CaseWorkbook, "TASKPANE_SNAPSHOT_CACHE_" + index.ToString("00", CultureInfo.InvariantCulture)));
            }

            return builder.ToString();
        }

        internal string GetCaseProperty(string propertyName)
        {
            return _excelInteropService.TryGetDocumentProperty(CaseWorkbook, propertyName);
        }

        public void Dispose()
        {
            foreach (Excel.Workbook workbook in Application.Workbooks.ToArray())
            {
                workbook.Close(false, null, null);
            }
        }
    }

    internal static class SnapshotLegacySerializer
    {
        private const string DefaultTabCaption = "その他";
        private const string AllTabCaption = "全て";

        internal static string Serialize(
            Excel.Workbook caseWorkbook,
            int masterVersion,
            Array values,
            IReadOnlyDictionary<int, long> fillColors,
            IReadOnlyDictionary<int, long> tabBackColors)
        {
            MasterTemplateSheetData sheetData = MasterTemplateSheetReader.BuildFromValues(
                values.GetUpperBound(0) + 2,
                values,
                rowIndex => fillColors.TryGetValue(rowIndex, out long color) ? color : 0L,
                rowIndex => tabBackColors.TryGetValue(rowIndex, out long color) ? color : 0L);

            List<string> lines = new List<string>
            {
                JoinFields(
                    "META",
                    TaskPaneSnapshotFormat.ExportVersion,
                    caseWorkbook.Name ?? string.Empty,
                    caseWorkbook.FullName ?? string.Empty,
                    BuildPreferredPaneWidth(sheetData).ToString(CultureInfo.InvariantCulture),
                    masterVersion.ToString(CultureInfo.InvariantCulture)),
                JoinFields(
                    "SPECIAL",
                    "btnCaseList",
                    GetCaseListCaption(caseWorkbook),
                    "caselist",
                    string.Empty,
                    "18",
                    "16",
                    "128",
                    "32",
                    GetCaseListBackColor(caseWorkbook).ToString(CultureInfo.InvariantCulture)),
                JoinFields(
                    "SPECIAL",
                    "btnAccounting",
                    "会計書類セット",
                    "accounting",
                    string.Empty,
                    "18",
                    "64",
                    "128",
                    "32",
                    "14348250")
            };

            Dictionary<string, int> tabOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, int> rowMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, long> tabColors = BuildTabBackColors(sheetData.Rows);

            foreach (MasterTemplateSheetRowData row in sheetData.Rows)
            {
                string key = row.Key;
                string caption = row.Caption;
                string templateFileName = row.TemplateFileName;
                string tabName = NormalizeTabName(row.TabName);
                if (key.Length == 0 || caption.Length == 0)
                {
                    continue;
                }

                if (!tabOrder.ContainsKey(tabName))
                {
                    int order = tabOrder.Count + 1;
                    tabOrder.Add(tabName, order);
                    rowMap[tabName] = 0;
                    long tabBackColor = tabColors.TryGetValue(tabName, out long value) ? value : 0L;
                    lines.Add(JoinFields("TAB", order.ToString(CultureInfo.InvariantCulture), tabName, tabBackColor.ToString(CultureInfo.InvariantCulture)));
                }

                rowMap[tabName]++;
                lines.Add(
                    JoinFields(
                        "DOC",
                        "btnDoc_" + key,
                        key,
                        caption,
                        "doc",
                        tabName,
                        rowMap[tabName].ToString(CultureInfo.InvariantCulture),
                        row.FillColor.ToString(CultureInfo.InvariantCulture),
                        templateFileName));
            }

            if (!tabOrder.ContainsKey(AllTabCaption))
            {
                lines.Add(JoinFields("TAB", (tabOrder.Count + 1).ToString(CultureInfo.InvariantCulture), AllTabCaption, "16777215"));
            }

            return string.Join("\r\n", lines);
        }

        private static Dictionary<string, long> BuildTabBackColors(IReadOnlyList<MasterTemplateSheetRowData> rows)
        {
            Dictionary<string, string> firstKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, long> tabColors = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            foreach (MasterTemplateSheetRowData row in rows ?? Array.Empty<MasterTemplateSheetRowData>())
            {
                string key = row.Key;
                string tabName = NormalizeTabName(row.TabName);
                if (key.Length == 0)
                {
                    continue;
                }

                if (!firstKeys.TryGetValue(tabName, out string currentKey) || CompareDocKeys(key, currentKey) < 0)
                {
                    firstKeys[tabName] = key;
                    tabColors[tabName] = row.TabBackColor;
                }
            }

            return tabColors;
        }

        private static int BuildPreferredPaneWidth(MasterTemplateSheetData sheetData)
        {
            int maxTabLength = 0;
            int maxCaptionLength = 0;
            foreach (MasterTemplateSheetRowData row in sheetData.Rows)
            {
                string tabName = NormalizeTabName(row.TabName);
                string caption = row.Caption ?? string.Empty;
                if (tabName.Length > maxTabLength)
                {
                    maxTabLength = tabName.Length;
                }

                if (caption.Length > maxCaptionLength)
                {
                    maxCaptionLength = caption.Length;
                }
            }

            int width = 80 + maxTabLength * 16 + maxCaptionLength * 12;
            if (width < 420)
            {
                return 420;
            }

            if (width > 900)
            {
                return 900;
            }

            return width;
        }

        private static string GetCaseListCaption(Excel.Workbook workbook)
        {
            return IsCaseListRegistered(workbook) ? "案件一覧登録（済）" : "案件一覧登録（未了）";
        }

        private static int GetCaseListBackColor(Excel.Workbook workbook)
        {
            return IsCaseListRegistered(workbook) ? 12566463 : 14803448;
        }

        private static bool IsCaseListRegistered(Excel.Workbook workbook)
        {
            if (workbook?.CustomDocumentProperties is not DocumentProperties properties)
            {
                return false;
            }

            return string.Equals(
                Convert.ToString(properties["CASELIST_REGISTERED"]?.Value, CultureInfo.InvariantCulture),
                "1",
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeTabName(string tabName)
        {
            string trimmed = (tabName ?? string.Empty).Trim();
            return trimmed.Length == 0 ? DefaultTabCaption : trimmed;
        }

        private static int CompareDocKeys(string leftKey, string rightKey)
        {
            if (long.TryParse(leftKey, out long left) && long.TryParse(rightKey, out long right))
            {
                return Math.Sign(left - right);
            }

            return string.Compare(leftKey, rightKey, StringComparison.OrdinalIgnoreCase);
        }

        private static string JoinFields(params string[] fields)
        {
            return string.Join("\t", fields.Select(EscapeField));
        }

        private static string EscapeField(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\t", "\\t")
                .Replace("\r\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\n", "\\n");
        }
    }

    internal sealed class SnapshotProjection
    {
        internal string ExportVersion { get; set; } = string.Empty;
        internal string MasterVersion { get; set; } = string.Empty;
        internal string WorkbookName { get; set; } = string.Empty;
        internal string WorkbookPath { get; set; } = string.Empty;
        internal int PreferredPaneWidth { get; set; }
        internal IReadOnlyList<string> SpecialButtons { get; set; } = Array.Empty<string>();
        internal IReadOnlyList<string> Tabs { get; set; } = Array.Empty<string>();
        internal IReadOnlyList<string> Docs { get; set; } = Array.Empty<string>();

        internal static SnapshotProjection FromSnapshotText(string snapshotText)
        {
            TaskPaneSnapshot snapshot = TaskPaneSnapshotParser.Parse(snapshotText);
            string[] lines = (snapshotText ?? string.Empty).Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            string[] metaFields = lines.Length == 0 ? Array.Empty<string>() : lines[0].Split('\t');

            return new SnapshotProjection
            {
                ExportVersion = metaFields.Length > 1 ? metaFields[1] : string.Empty,
                MasterVersion = metaFields.Length > 5 ? metaFields[5] : string.Empty,
                WorkbookName = snapshot.WorkbookName ?? string.Empty,
                WorkbookPath = snapshot.WorkbookPath ?? string.Empty,
                PreferredPaneWidth = snapshot.PreferredPaneWidth,
                SpecialButtons = snapshot.SpecialButtons
                    .Select(button => string.Join("|",
                        button.Name ?? string.Empty,
                        button.Caption ?? string.Empty,
                        button.ActionKind ?? string.Empty,
                        button.Key ?? string.Empty,
                        button.Left.ToString(CultureInfo.InvariantCulture),
                        button.Top.ToString(CultureInfo.InvariantCulture),
                        button.Width.ToString(CultureInfo.InvariantCulture),
                        button.Height.ToString(CultureInfo.InvariantCulture),
                        ToOleColor(button.BackColor).ToString(CultureInfo.InvariantCulture)))
                    .ToArray(),
                Tabs = snapshot.Tabs
                    .Select(tab => string.Join("|",
                        tab.Order.ToString(CultureInfo.InvariantCulture),
                        tab.TabName ?? string.Empty,
                        ToOleColor(tab.BackColor).ToString(CultureInfo.InvariantCulture)))
                    .ToArray(),
                Docs = snapshot.DocButtons
                    .Select(doc => string.Join("|",
                        doc.Name ?? string.Empty,
                        doc.Key ?? string.Empty,
                        doc.Caption ?? string.Empty,
                        doc.ActionKind ?? string.Empty,
                        doc.TabName ?? string.Empty,
                        doc.RowIndex.ToString(CultureInfo.InvariantCulture),
                        ToOleColor(doc.FillColor).ToString(CultureInfo.InvariantCulture),
                        doc.TemplateFileName ?? string.Empty))
                    .ToArray()
            };
        }

        internal string ToJson()
        {
            return "{"
                + "\"exportVersion\":\"" + EscapeJson(ExportVersion) + "\","
                + "\"masterVersion\":\"" + EscapeJson(MasterVersion) + "\","
                + "\"workbookName\":\"" + EscapeJson(WorkbookName) + "\","
                + "\"workbookPath\":\"" + EscapeJson(WorkbookPath) + "\","
                + "\"preferredPaneWidth\":" + PreferredPaneWidth.ToString(CultureInfo.InvariantCulture) + ","
                + "\"specialButtons\":" + ToJsonArray(SpecialButtons) + ","
                + "\"tabs\":" + ToJsonArray(Tabs) + ","
                + "\"docs\":" + ToJsonArray(Docs)
                + "}";
        }

        private static int ToOleColor(Color color)
        {
            return color.IsEmpty ? 0 : ColorTranslator.ToOle(color);
        }

        private static string ToJsonArray(IEnumerable<string> values)
        {
            return "[" + string.Join(",", (values ?? Array.Empty<string>()).Select(value => "\"" + EscapeJson(value ?? string.Empty) + "\"")) + "]";
        }

        private static string EscapeJson(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }
    }
}
