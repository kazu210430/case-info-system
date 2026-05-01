using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class TaskPaneSnapshotBuilderService : CaseInfoSystem.ExcelAddIn.App.ICaseTaskPaneSnapshotReader
	{
		internal sealed class TaskPaneBuildResult
		{
			internal string SnapshotText { get; private set; }

			internal bool UpdatedCaseSnapshotCache { get; private set; }

			internal TaskPaneBuildResult (string snapshotText, bool updatedCaseSnapshotCache)
			{
				SnapshotText = snapshotText ?? string.Empty;
				UpdatedCaseSnapshotCache = updatedCaseSnapshotCache;
			}
		}

		private const string LineMeta = "META";

		private const string LineSpecial = "SPECIAL";

		private const string LineTab = "TAB";

		private const string LineDoc = "DOC";

		private const string TaskPaneCacheCountProp = "TASKPANE_SNAPSHOT_CACHE_COUNT";

		private const string TaskPaneCachePartPropPrefix = "TASKPANE_SNAPSHOT_CACHE_";

		private const string TaskPaneBaseCacheCountProp = "TASKPANE_BASE_SNAPSHOT_COUNT";

		private const string TaskPaneBaseCachePartPropPrefix = "TASKPANE_BASE_SNAPSHOT_";

		private const string TaskPaneBaseMasterVersionProp = "TASKPANE_BASE_MASTER_VERSION";

		private const string TaskPaneMasterVersionProp = "TASKPANE_MASTER_VERSION";

		private const string MasterSheetName = "雛形一覧";

		private const string MasterSheetCodeName = "shMasterList";

		private const int MasterListFirstDataRow = 3;

		private const int CaseListButtonBackColorUnregistered = 14803448;

		private const int CaseListButtonBackColorRegistered = 12566463;

		private const int AccountingButtonBackColor = 14348250;

		private const int DefaultAllTabBackColor = 16777215;

		private const string SystemRootPropertyName = "SYSTEM_ROOT";

		private static readonly XNamespace CustomPropertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

		private readonly Application _application;

		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly IMasterTemplateSheetReader _masterTemplateSheetReader;

		private readonly Logger _logger;

		internal TaskPaneSnapshotBuilderService (Application application, ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService, IMasterTemplateSheetReader masterTemplateSheetReader, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_masterTemplateSheetReader = masterTemplateSheetReader ?? throw new ArgumentNullException ("masterTemplateSheetReader");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		public TaskPaneBuildResult BuildSnapshotText (Workbook workbook)
		{
			if (workbook == null) {
				return new TaskPaneBuildResult (string.Empty, updatedCaseSnapshotCache: false);
			}
			long num = 0L;
			bool openedNow = false;
			Workbook workbook2 = null;
			try {
				long masterVersion = 0L;
				num = 5L;
				string text = LoadSnapshotCache (workbook, "TASKPANE_SNAPSHOT_CACHE_COUNT", "TASKPANE_SNAPSHOT_CACHE_");
				if (!string.IsNullOrWhiteSpace (text)) {
					num = 8L;
					if (!TaskPaneSnapshotFormat.IsCompatible (text)) {
						ClearSnapshotCache (workbook, "TASKPANE_SNAPSHOT_CACHE_COUNT", "TASKPANE_SNAPSHOT_CACHE_");
						_logger.Info ("Task pane snapshot incompatible CASE cache was cleared. exportVersion=" + TaskPaneSnapshotFormat.TryReadExportVersion (text));
						text = string.Empty;
					}
					if (!string.IsNullOrWhiteSpace (text) && TryReadLatestMasterVersion (workbook, out masterVersion)) {
						long documentPropertyLong = GetDocumentPropertyLong (workbook, "TASKPANE_MASTER_VERSION", 0L);
						if (masterVersion <= 0 || masterVersion <= documentPropertyLong) {
							string snapshotText2 = ApplyDynamicSpecialButtonOverrides (text, workbook);
							_logger.Info ("Task pane snapshot source=CaseCache, caseListCaption=" + GetCaseListCaption (workbook) + ", cacheCount=" + (_excelInteropService.TryGetDocumentProperty (workbook, "TASKPANE_SNAPSHOT_CACHE_COUNT") ?? string.Empty));
							return new TaskPaneBuildResult (snapshotText2, updatedCaseSnapshotCache: false);
						}
						_logger.Info ("Task pane snapshot case cache is stale. caseMasterVersion=" + documentPropertyLong + ", latestMasterVersion=" + masterVersion);
					}
				}
				num = 10L;
				string text2 = LoadSnapshotCache (workbook, "TASKPANE_BASE_SNAPSHOT_COUNT", "TASKPANE_BASE_SNAPSHOT_");
				if (!string.IsNullOrWhiteSpace (text2)) {
					num = 12L;
					if (!TaskPaneSnapshotFormat.IsCompatible (text2)) {
						ClearSnapshotCache (workbook, "TASKPANE_BASE_SNAPSHOT_COUNT", "TASKPANE_BASE_SNAPSHOT_");
						_logger.Info ("Task pane snapshot incompatible Base cache was cleared. exportVersion=" + TaskPaneSnapshotFormat.TryReadExportVersion (text2));
						text2 = string.Empty;
					}
					if (!string.IsNullOrWhiteSpace (text2)) {
						string text3 = _excelInteropService.TryGetDocumentProperty (workbook, "TASKPANE_BASE_MASTER_VERSION") ?? string.Empty;
						long result = 0L;
						long.TryParse (text3, out result);
						if (!TryReadLatestMasterVersion (workbook, out masterVersion)) {
							string snapshotText3 = ApplyDynamicSpecialButtonOverrides (text2, workbook);
							SaveCaseSnapshotCache (workbook, snapshotText3);
							if (!string.IsNullOrWhiteSpace (text3)) {
								_excelInteropService.SetDocumentProperty (workbook, "TASKPANE_MASTER_VERSION", text3);
							}
							_logger.Info ("Task pane snapshot source=BaseCacheFallback, caseListCaption=" + GetCaseListCaption (workbook) + ", baseCacheCount=" + (_excelInteropService.TryGetDocumentProperty (workbook, "TASKPANE_BASE_SNAPSHOT_COUNT") ?? string.Empty));
							return new TaskPaneBuildResult (snapshotText3, updatedCaseSnapshotCache: true);
						}
						if (masterVersion <= 0 || masterVersion <= result) {
							string snapshotText4 = ApplyDynamicSpecialButtonOverrides (text2, workbook);
							SaveCaseSnapshotCache (workbook, snapshotText4);
							if (!string.IsNullOrWhiteSpace (text3)) {
								_excelInteropService.SetDocumentProperty (workbook, "TASKPANE_MASTER_VERSION", text3);
							}
							_logger.Info ("Task pane snapshot source=BaseCache, caseListCaption=" + GetCaseListCaption (workbook) + ", baseCacheCount=" + (_excelInteropService.TryGetDocumentProperty (workbook, "TASKPANE_BASE_SNAPSHOT_COUNT") ?? string.Empty));
							return new TaskPaneBuildResult (snapshotText4, updatedCaseSnapshotCache: true);
						}
						_logger.Info ("Task pane snapshot base cache is stale. embeddedMasterVersion=" + result + ", latestMasterVersion=" + masterVersion);
					}
				}
				num = 20L;
				if (workbook2 == null) {
					workbook2 = OpenMasterReadOnly (workbook, out openedNow);
				}
				Worksheet masterListWorksheet = GetMasterListWorksheet (workbook2);
				if (masterListWorksheet == null) {
					throw new InvalidOperationException ("雛形一覧シートが見つかりません。");
				}
				num = 30L;
				long documentPropertyLong2 = GetDocumentPropertyLong (workbook2, "TASKPANE_MASTER_VERSION", 0L);
				_excelInteropService.SetDocumentProperty (workbook, "TASKPANE_MASTER_VERSION", documentPropertyLong2.ToString ());
				List<string> list = new List<string> ();
				Dictionary<string, int> tabOrder = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
				Dictionary<string, int> rowMap = new Dictionary<string, int> (StringComparer.OrdinalIgnoreCase);
				num = 40L;
				list.Add (JoinFields ("META", "2", _excelInteropService.GetWorkbookName (workbook), _excelInteropService.GetWorkbookFullName (workbook), BuildPreferredPaneWidthFromMasterSheet (masterListWorksheet).ToString (), documentPropertyLong2.ToString ()));
				num = 50L;
				AppendSpecialButtons (list, workbook);
				num = 60L;
				AppendTemplateDefinitions (list, tabOrder, rowMap, masterListWorksheet);
				num = 70L;
				string snapshotText5 = string.Join ("\r\n", list);
				SaveCaseSnapshotCache (workbook, snapshotText5);
				_logger.Info ("Task pane snapshot source=MasterListRebuild, caseListCaption=" + GetCaseListCaption (workbook) + ", masterVersion=" + documentPropertyLong2);
				return new TaskPaneBuildResult (snapshotText5, updatedCaseSnapshotCache: true);
			} catch (Exception ex) {
				_logger.Error ("TaskPaneSnapshotBuilderService.BuildSnapshotText failed. step=" + num, ex);
				return new TaskPaneBuildResult (JoinFields ("META", "2", "ERROR", "step=" + num + " / " + ex.Message), updatedCaseSnapshotCache: false);
			} finally {
				if (openedNow && workbook2 != null) {
					try {
						workbook2.Close (false, Type.Missing, Type.Missing);
					} catch (Exception exception) {
						_logger.Error ("TaskPaneSnapshotBuilderService.BuildSnapshotText close failed.", exception);
					}
				}
			}
		}


		private bool TryReadLatestMasterVersion (Workbook caseWorkbook, out long masterVersion)
		{
			masterVersion = 0L;
			if (caseWorkbook == null) {
				return false;
			}
			try {
				string text = ResolveMasterPath (caseWorkbook);
				if (string.IsNullOrWhiteSpace (text) || !_pathCompatibilityService.FileExistsSafe (text)) {
					return false;
				}
				Workbook workbook = FindOpenMasterWorkbook (text);
				if (workbook != null) {
					masterVersion = GetDocumentPropertyLong (workbook, "TASKPANE_MASTER_VERSION", 0L);
					return true;
				}
				return TryReadDocumentPropertyFromPackage (text, "TASKPANE_MASTER_VERSION", out masterVersion);
			} catch (Exception exception) {
				_logger.Error ("TaskPaneSnapshotBuilderService.TryReadLatestMasterVersion failed.", exception);
				return false;
			}
		}

		private bool TryReadDocumentPropertyFromPackage (string workbookPath, string propertyName, out long propertyValue)
		{
			propertyValue = 0L;
			if (string.IsNullOrWhiteSpace (workbookPath) || string.IsNullOrWhiteSpace (propertyName)) {
				return false;
			}
			try {
				ZipArchive val = ZipFile.OpenRead (workbookPath);
				try {
					ZipArchiveEntry entry = val.GetEntry ("docProps/custom.xml");
					if (entry == null) {
						return false;
					}
					using (Stream stream = entry.Open ()) {
						XDocument xDocument = XDocument.Load (stream);
						IEnumerable<XElement> enumerable;
						if (xDocument.Root != null) {
							enumerable = xDocument.Root.Elements (CustomPropertiesNamespace + "property");
						} else {
							IEnumerable<XElement> enumerable2 = Array.Empty<XElement> ();
							enumerable = enumerable2;
						}
						foreach (XElement item in enumerable) {
							XAttribute xAttribute = item.Attribute ("name");
							if (xAttribute == null || !string.Equals (xAttribute.Value, propertyName, StringComparison.OrdinalIgnoreCase)) {
								continue;
							}
							XElement xElement = item.Elements ().FirstOrDefault ();
							if (xElement == null) {
								return false;
							}
							return long.TryParse (xElement.Value ?? string.Empty, out propertyValue);
						}
					}
				} finally {
					((IDisposable)val)?.Dispose ();
				}
			} catch (Exception exception) {
				_logger.Error ("TaskPaneSnapshotBuilderService.TryReadDocumentPropertyFromPackage failed. workbookPath=" + workbookPath + ", propertyName=" + propertyName, exception);
			}
			return false;
		}

		private void AppendSpecialButtons (List<string> lines, Workbook workbook)
		{
			string caseListCaption = GetCaseListCaption (workbook);
			string text = GetCaseListBackColor (workbook).ToString ();
			lines.Add (JoinFields ("SPECIAL", "btnCaseList", caseListCaption, "caselist", string.Empty, "18", "16", "128", "32", text));
			lines.Add (JoinFields ("SPECIAL", "btnAccounting", "会計書類セット", "accounting", string.Empty, "18", "64", "128", "32", 14348250.ToString ()));
		}

		private void AppendTemplateDefinitions (List<string> lines, Dictionary<string, int> tabOrder, Dictionary<string, int> rowMap, Worksheet masterWorksheet)
		{
			MasterTemplateSheetData masterSheetData = _masterTemplateSheetReader.Read (masterWorksheet);
			if (masterSheetData.LastRow < 3) {
				return;
			}
			Dictionary<string, long> tabBackColors = BuildTabBackColors (masterSheetData.Rows, normalizeBlankTabName: false);
			string text = "全て";
			foreach (MasterTemplateSheetRowData row in masterSheetData.Rows) {
				string key = row.Key;
				string templateFile = row.TemplateFileName;
				string caption = row.Caption;
				string text2 = row.TabName;
				if (!string.IsNullOrWhiteSpace (key) && !string.IsNullOrWhiteSpace (caption)) {
					if (string.IsNullOrWhiteSpace (text2)) {
						text2 = "その他";
					}
					if (!tabOrder.ContainsKey (text2)) {
						int value = tabOrder.Count + 1;
						tabOrder.Add (text2, value);
						lines.Add (JoinFields ("TAB", value.ToString (), text2, GetTabBackColor (tabBackColors, text2).ToString ()));
						rowMap.Add (text2, 0);
					}
					rowMap [text2]++;
					long fillColor = row.FillColor;
					lines.Add (JoinFields ("DOC", "btnDoc_" + key, key, caption, "doc", text2, rowMap [text2].ToString (), fillColor.ToString (), templateFile));
				}
			}
			if (!tabOrder.ContainsKey (text)) {
				int value2 = tabOrder.Count + 1;
				tabOrder.Add (text, value2);
				lines.Add (JoinFields ("TAB", value2.ToString (), text, 16777215.ToString ()));
			}
		}

		private Workbook OpenMasterReadOnly (Workbook caseWorkbook, out bool openedNow)
		{
			openedNow = false;
			string text = ResolveMasterPath (caseWorkbook);
			if (string.IsNullOrWhiteSpace (text)) {
				throw new InvalidOperationException ("Masterブックのパスを解決できませんでした。");
			}
			Workbook workbook = FindOpenMasterWorkbook (text);
			if (workbook != null) {
				HideWorkbookWindows (workbook);
				return workbook;
			}
			if (!_pathCompatibilityService.FileExistsSafe (text)) {
				throw new FileNotFoundException ("Masterブックが見つかりません。", text);
			}
			bool enableEvents = _application.EnableEvents;
			try {
				_application.EnableEvents = false;
				Workbook workbook2 = _application.Workbooks.Open (text, 0, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing);
				HideWorkbookWindows (workbook2);
				openedNow = true;
				return workbook2;
			} finally {
				_application.EnableEvents = enableEvents;
			}
		}

		private Workbook FindOpenMasterWorkbook (string masterPath)
		{
			if (string.IsNullOrWhiteSpace (masterPath)) {
				return null;
			}
			Workbook workbook = _excelInteropService.FindOpenWorkbook (masterPath);
			if (workbook != null) {
				return workbook;
			}
			string fileNameFromPath = _pathCompatibilityService.GetFileNameFromPath (masterPath);
			if (string.IsNullOrWhiteSpace (fileNameFromPath)) {
				return null;
			}
			foreach (Workbook workbook2 in _application.Workbooks) {
				string workbookName = _excelInteropService.GetWorkbookName (workbook2);
				if (string.Equals (workbookName, fileNameFromPath, StringComparison.OrdinalIgnoreCase)) {
					return workbook2;
				}
			}
			return null;
		}

		private string ResolveMasterPath (Workbook caseWorkbook)
		{
			string text = _pathCompatibilityService.NormalizePath (_excelInteropService.TryGetDocumentProperty (caseWorkbook, "SYSTEM_ROOT"));
			if (text.Length > 0) {
				return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath (text, _pathCompatibilityService);
			}
			string text2 = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookPath (caseWorkbook));
			if (text2.Length > 0) {
				string text3 = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath (text2, _pathCompatibilityService);
				if (_pathCompatibilityService.FileExistsSafe (text3)) {
					return text3;
				}
				string text4 = _pathCompatibilityService.NormalizePath (_pathCompatibilityService.GetParentFolderPath (text2));
				if (text4.Length > 0) {
					return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath (text4, _pathCompatibilityService);
				}
			}
			return string.Empty;
		}

		private Worksheet GetMasterListWorksheet (Workbook masterWorkbook)
		{
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (masterWorkbook, "shMasterList");
			if (worksheet != null) {
				return worksheet;
			}
			try {
				return masterWorkbook.Worksheets ["雛形一覧"] as Worksheet;
			} catch (Exception exception) {
				_logger.Error ("TaskPaneSnapshotBuilderService.GetMasterListWorksheet failed.", exception);
				return null;
			}
		}

		private static void HideWorkbookWindows (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			Windows windows = null;
			try {
				windows = workbook.Windows;
				int windowCount = (windows == null) ? 0 : windows.Count;
				for (int index = 1; index <= windowCount; index++) {
					Window window = null;
					try {
						window = windows [index];
						if (window != null) {
							window.Visible = false;
						}
					} finally {
						if (window != null && Marshal.IsComObject (window)) {
							ComObjectReleaseService.Release (window);
						}
					}
				}
			} catch {
			} finally {
				if (windows != null && Marshal.IsComObject (windows)) {
					ComObjectReleaseService.Release (windows);
				}
			}
		}

		private static int CompareDocKeys (string leftKey, string rightKey)
		{
			if (long.TryParse (leftKey, out var result) && long.TryParse (rightKey, out var result2)) {
				return Math.Sign (result - result2);
			}
			return string.Compare (leftKey, rightKey, StringComparison.OrdinalIgnoreCase);
		}

		private int BuildPreferredPaneWidthFromMasterSheet (Worksheet masterWorksheet)
		{
			if (masterWorksheet == null) {
				return 720;
			}
			try {
				MasterTemplateSheetData masterSheetSnapshot = _masterTemplateSheetReader.Read (masterWorksheet);
				int num = 0;
				int num2 = 0;
				for (int i = 0; i < masterSheetSnapshot.Rows.Count; i++) {
					string tabName = masterSheetSnapshot.Rows [i].TabName;
					string caption = masterSheetSnapshot.Rows [i].Caption;
					if (tabName.Length > num) {
						num = tabName.Length;
					}
					if (caption.Length > num2) {
						num2 = caption.Length;
					}
				}
				int num3 = 80 + num * 16 + num2 * 12;
				if (num3 < 420) {
					num3 = 420;
				}
				if (num3 > 900) {
					num3 = 900;
				}
				return num3;
			} catch {
				return 720;
			}
		}

		private static Dictionary<string, long> BuildTabBackColors (IReadOnlyList<MasterTemplateSheetRowData> rows, bool normalizeBlankTabName)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			Dictionary<string, long> dictionary2 = new Dictionary<string, long> (StringComparer.OrdinalIgnoreCase);
			if (rows == null) {
				return dictionary2;
			}
			for (int i = 0; i < rows.Count; i++) {
				MasterTemplateSheetRowData masterTemplateSheetRowData = rows [i];
				string key = masterTemplateSheetRowData.Key;
				string text = masterTemplateSheetRowData.TabName;
				if (normalizeBlankTabName && text.Length == 0) {
					text = "その他";
				}
				if (key.Length != 0 && (!dictionary.TryGetValue (text, out var value) || CompareDocKeys (key, value) < 0)) {
					dictionary [text] = key;
					dictionary2 [text] = masterTemplateSheetRowData.TabBackColor;
				}
			}
			return dictionary2;
		}

		private static long GetTabBackColor (IReadOnlyDictionary<string, long> tabBackColors, string tabName)
		{
			if (tabBackColors == null || string.IsNullOrWhiteSpace (tabName)) {
				return 0L;
			}
			long value;
			return tabBackColors.TryGetValue (tabName, out value) ? value : 0L;
		}

		private static void ReleaseComObject (object comObject)
		{
			// Snapshot 構築中に所有した COM 参照は完全解放の方針を維持する。
			ComObjectReleaseService.FinalRelease (comObject);
		}

		private string LoadSnapshotCache (Workbook workbook, string countPropName, string partPropPrefix)
		{
			return TaskPaneSnapshotChunkReadHelper.LoadSnapshot (_excelInteropService, workbook, countPropName, partPropPrefix);
		}

		private void SaveCaseSnapshotCache (Workbook workbook, string snapshotText)
		{
			TaskPaneSnapshotChunkStorageHelper.SaveSnapshot (
				_excelInteropService,
				workbook,
				TaskPaneCacheCountProp,
				TaskPaneCachePartPropPrefix,
				snapshotText);
		}

		private void ClearSnapshotCache (Workbook workbook, string countPropName, string partPropPrefix)
		{
			TaskPaneSnapshotChunkStorageHelper.ClearSnapshot (_excelInteropService, workbook, countPropName, partPropPrefix);
		}

		private string ApplyDynamicSpecialButtonOverrides (string snapshotText, Workbook workbook)
		{
			if (string.IsNullOrWhiteSpace (snapshotText) || workbook == null) {
				return snapshotText ?? string.Empty;
			}
			string caseListCaption = GetCaseListCaption (workbook);
			string text = GetCaseListBackColor (workbook).ToString ();
			string text2 = 14348250.ToString ();
			string[] array = snapshotText.Replace ("\r\n", "\n").Split ('\n');
			bool flag = false;
			bool flag2 = false;
			for (int i = 0; i < array.Length; i++) {
				string text3 = array [i] ?? string.Empty;
				if (!text3.StartsWith ("SPECIAL\t", StringComparison.Ordinal)) {
					continue;
				}
				string[] array2 = text3.Split ('\t');
				if (array2.Length >= 10) {
					if (string.Equals (array2 [1], "btnCaseList", StringComparison.OrdinalIgnoreCase)) {
						array2 [2] = caseListCaption;
						array2 [9] = text;
						array [i] = JoinFields (array2);
						flag = true;
					} else if (string.Equals (array2 [1], "btnAccounting", StringComparison.OrdinalIgnoreCase)) {
						array2 [9] = text2;
						array [i] = JoinFields (array2);
						flag2 = true;
					}
				}
			}
			if (!flag || !flag2) {
				List<string> list = new List<string> (array.Length + 2);
				bool flag3 = false;
				string[] array3 = array;
				foreach (string text4 in array3) {
					list.Add (text4);
					if (!flag3 && text4.StartsWith ("META\t", StringComparison.Ordinal)) {
						flag3 = true;
						if (!flag) {
							list.Add (JoinFields ("SPECIAL", "btnCaseList", caseListCaption, "caselist", string.Empty, "18", "16", "128", "32", text));
						}
						if (!flag2) {
							list.Add (JoinFields ("SPECIAL", "btnAccounting", "会計書類セット", "accounting", string.Empty, "18", "64", "128", "32", text2));
						}
					}
				}
				array = list.ToArray ();
			}
			return string.Join ("\r\n", array);
		}

		private static string JoinFields (params string[] values)
		{
			if (values == null || values.Length == 0) {
				return string.Empty;
			}
			string[] array = new string[values.Length];
			for (int i = 0; i < values.Length; i++) {
				array [i] = EscapeField (values [i] ?? string.Empty);
			}
			return string.Join ("\t", array);
		}

		private static string EscapeField (string value)
		{
			return (value ?? string.Empty).Replace ("\\", "\\\\").Replace ("\t", "\\t").Replace ("\r\n", "\\n")
				.Replace ("\r", "\\n")
				.Replace ("\n", "\\n");
		}

		private static string NormalizeDocKey (string key)
		{
			string text = (key ?? string.Empty).Trim ();
			if (text.Length == 0) {
				return string.Empty;
			}
			long result;
			return long.TryParse (text, out result) ? result.ToString ("00") : text;
		}

		private long GetDocumentPropertyLong (Workbook workbook, string propName, long defaultValue)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (propName)) {
				return defaultValue;
			}
			try {
				if (!(workbook.CustomDocumentProperties is DocumentProperties documentProperties)) {
					return defaultValue;
				}
				DocumentProperty documentProperty = documentProperties [propName];
				if (documentProperty == null) {
					return defaultValue;
				}
				long num = default(long);
				return (long.TryParse (Convert.ToString ((dynamic)documentProperty.Value), out num)) ? num : defaultValue;
			} catch {
				return defaultValue;
			}
		}

		private string GetCaseListCaption (Workbook workbook)
		{
			string a = _excelInteropService.TryGetDocumentProperty (workbook, "CASELIST_REGISTERED");
			return string.Equals (a, "1", StringComparison.OrdinalIgnoreCase) ? "案件一覧登録（済）" : "案件一覧登録（未了）";
		}

		private int GetCaseListBackColor (Workbook workbook)
		{
			string a = _excelInteropService.TryGetDocumentProperty (workbook, "CASELIST_REGISTERED");
			return string.Equals (a, "1", StringComparison.OrdinalIgnoreCase) ? 12566463 : 14803448;
		}
	}
}
