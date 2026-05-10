using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class BaseHomeFieldInventorySyncService
	{
		internal sealed class SyncResult
		{
			internal bool Success { get; set; }

			internal int CheckedCount { get; set; }

			internal int UpdatedCount { get; set; }

			internal int UnchangedCount { get; set; }

			internal int WarningCount { get; set; }

			internal string Message { get; set; }
		}

		internal sealed class BaseHomeFieldKeyRow
		{
			internal int RowNumber { get; set; }

			internal string FieldKey { get; set; }
		}

		internal sealed class FieldInventoryRow
		{
			internal int RowNumber { get; set; }

			internal string SourceCell { get; set; }

			internal string ProposedFieldKey { get; set; }
		}

		internal sealed class FieldInventoryUpdate
		{
			internal int BaseHomeRowNumber { get; set; }

			internal int FieldInventoryRowNumber { get; set; }

			internal string OldFieldKey { get; set; }

			internal string NewFieldKey { get; set; }
		}

		internal sealed class SyncPlan
		{
			internal List<FieldInventoryUpdate> Updates { get; } = new List<FieldInventoryUpdate>();

			internal List<string> Errors { get; } = new List<string>();

			internal List<string> Warnings { get; } = new List<string>();

			internal int CheckedCount { get; set; }

			internal int UnchangedCount { get; set; }

			internal bool CanApply => Errors.Count == 0;
		}

		private sealed class SheetProtectionState
		{
			internal bool IsProtected { get; set; }

			internal bool AllowFormattingCells { get; set; }

			internal bool AllowFormattingColumns { get; set; }

			internal bool AllowFormattingRows { get; set; }

			internal bool AllowInsertingColumns { get; set; }

			internal bool AllowInsertingRows { get; set; }

			internal bool AllowInsertingHyperlinks { get; set; }

			internal bool AllowDeletingColumns { get; set; }

			internal bool AllowDeletingRows { get; set; }

			internal bool AllowSorting { get; set; }

			internal bool AllowFiltering { get; set; }

			internal bool AllowUsingPivotTables { get; set; }
		}

		private sealed class TemporarySheetProtectionRestoreScope : IDisposable
		{
			private readonly Excel.Worksheet _worksheet;

			private readonly SheetProtectionState _state;

			private bool _disposed;

			internal TemporarySheetProtectionRestoreScope(Excel.Worksheet worksheet)
			{
				_worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
				_state = SaveSheetProtectionState(worksheet);
				if (_state.IsProtected)
				{
					worksheet.Unprotect(string.Empty);
				}
			}

			public void Dispose()
			{
				if (_disposed)
				{
					return;
				}

				_disposed = true;
				RestoreSheetProtectionState(_worksheet, _state);
			}
		}

		private const string BaseHomeSheetCodeName = "shHOME";
		private const string BaseHomeSheetName = "ホーム";
		private const string FieldInventorySheetName = "CaseList_FieldInventory";
		private const string SourceCellHeaderName = "SourceCell";
		private const string ProposedFieldKeyHeaderName = "ProposedFieldKey";
		private const string SystemRootPropertyName = "SYSTEM_ROOT";

		private static readonly string[] ImportantFieldKeys =
		{
			"顧客_名前",
			"顧客_よみ",
			"顧客_敬称",
			FieldKeyRenameMap.LegacyLawyerKey,
			FieldKeyRenameMap.CurrentLawyerKey,
			FieldKeyRenameMap.LegacyPostalCodeKey,
			FieldKeyRenameMap.CurrentPostalCodeKey,
			FieldKeyRenameMap.LegacyAddressKey,
			FieldKeyRenameMap.CurrentAddressKey,
			FieldKeyRenameMap.LegacyOfficeNameKey,
			FieldKeyRenameMap.CurrentOfficeNameKey,
			FieldKeyRenameMap.LegacyPhoneKey,
			FieldKeyRenameMap.CurrentPhoneKey,
			FieldKeyRenameMap.LegacyFaxKey,
			FieldKeyRenameMap.CurrentFaxKey
		};

		private readonly Excel.Application _application;

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

		private readonly Logger _logger;

		internal BaseHomeFieldInventorySyncService(
			Excel.Application application,
			KernelWorkbookService kernelWorkbookService,
			ExcelInteropService excelInteropService,
			PathCompatibilityService pathCompatibilityService,
			CaseWorkbookLifecycleService caseWorkbookLifecycleService,
			Logger logger)
		{
			_application = application ?? throw new ArgumentNullException(nameof(application));
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException(nameof(kernelWorkbookService));
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
			_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException(nameof(caseWorkbookLifecycleService));
			_logger = logger ?? throw new ArgumentNullException(nameof(logger));
		}

		internal SyncResult Execute(WorkbookContext context)
		{
			if (context == null)
			{
				throw new InvalidOperationException("WorkbookContext is required for Base HOME field inventory sync.");
			}

			Excel.Workbook kernelWorkbook = _kernelWorkbookService.ResolveKernelWorkbook(context);
			if (kernelWorkbook == null)
			{
				throw new InvalidOperationException("Kernel ブックを開いてから実行してください。");
			}

			string systemRoot = ResolveSystemRoot(context, kernelWorkbook);
			string baseWorkbookPath = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath(systemRoot, _pathCompatibilityService);
			if (!_pathCompatibilityService.FileExistsSafe(baseWorkbookPath))
			{
				throw new InvalidOperationException("Base が見つかりません: " + baseWorkbookPath);
			}

			Excel.Workbook baseWorkbook = _excelInteropService.FindOpenWorkbook(baseWorkbookPath);
			bool openedBaseWorkbook = baseWorkbook == null;
			try
			{
				using (var applicationStateScope = new ExcelApplicationStateScope(_application, suppressRestoreExceptions: true))
				{
					applicationStateScope.SetScreenUpdating(false);
					applicationStateScope.SetEnableEvents(false);
					applicationStateScope.SetDisplayAlerts(false);

					if (baseWorkbook == null)
					{
						baseWorkbook = _application.Workbooks.Open(baseWorkbookPath, 0, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
					}

					Excel.Worksheet baseHomeWorksheet = ResolveBaseHomeWorksheet(baseWorkbook);
					Excel.Worksheet fieldInventoryWorksheet = _excelInteropService.FindWorksheetByName(kernelWorkbook, FieldInventorySheetName);
					if (baseHomeWorksheet == null)
					{
						throw new InvalidOperationException("Base の ホーム シートを取得できませんでした。");
					}

					if (fieldInventoryWorksheet == null)
					{
						throw new InvalidOperationException("Kernelブックの管理シート CaseList_FieldInventory を読み取れません。");
					}

					IReadOnlyList<BaseHomeFieldKeyRow> baseRows = ReadBaseHomeFieldKeyRows(baseHomeWorksheet);
					int proposedFieldKeyColumnIndex;
					IReadOnlyList<FieldInventoryRow> inventoryRows = ReadFieldInventoryRows(fieldInventoryWorksheet, out proposedFieldKeyColumnIndex);
					SyncPlan plan = BuildPlan(baseRows, inventoryRows, ImportantFieldKeys);
					if (!plan.CanApply)
					{
						return BuildFailureResult(plan);
					}

					if (plan.Updates.Count > 0)
					{
						using (new TemporarySheetProtectionRestoreScope(fieldInventoryWorksheet))
						{
							ApplyUpdates(fieldInventoryWorksheet, proposedFieldKeyColumnIndex, plan.Updates);
						}

						kernelWorkbook.Save();
					}

					_logger.Info("Base HOME field inventory sync completed. checked=" + plan.CheckedCount.ToString() + ", updated=" + plan.Updates.Count.ToString() + ", unchanged=" + plan.UnchangedCount.ToString() + ", warnings=" + plan.Warnings.Count.ToString());
					return BuildSuccessResult(plan);
				}
			}
			finally
			{
				if (openedBaseWorkbook && baseWorkbook != null)
				{
					try
					{
						using (var closeScope = new ExcelApplicationStateScope(_application, suppressRestoreExceptions: true))
						{
							closeScope.SetDisplayAlerts(false);
							using (_caseWorkbookLifecycleService.BeginManagedCloseScope(baseWorkbook))
							{
								baseWorkbook.Close(false, Type.Missing, Type.Missing);
							}
						}
					}
					catch (Exception exception)
					{
						_logger.Warn("Base HOME field inventory sync could not close opened Base workbook: " + exception.Message);
					}
				}
			}
		}

		internal static SyncPlan BuildPlan(
			IReadOnlyList<BaseHomeFieldKeyRow> baseRows,
			IReadOnlyList<FieldInventoryRow> inventoryRows,
			IReadOnlyCollection<string> importantFieldKeys)
		{
			var plan = new SyncPlan();
			var importantKeys = new HashSet<string>(importantFieldKeys ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
			var seenBaseKeys = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
			var seenBaseRows = new HashSet<int>();
			var inventoryByBaseRow = new Dictionary<int, FieldInventoryRow>();
			var finalKeysByInventoryRow = new Dictionary<int, string>();

			if (baseRows == null || baseRows.Count == 0)
			{
				plan.Errors.Add("Base HOME A列のフィールドキーを読み取れませんでした。");
				return plan;
			}

			if (inventoryRows == null || inventoryRows.Count == 0)
			{
				plan.Errors.Add("CaseList_FieldInventory の既存定義を読み取れませんでした。");
				return plan;
			}

			foreach (FieldInventoryRow inventoryRow in inventoryRows)
			{
				if (inventoryRow == null)
				{
					continue;
				}

				finalKeysByInventoryRow[inventoryRow.RowNumber] = NormalizeKey(inventoryRow.ProposedFieldKey);
				int baseRowNumber;
				if (!TryParseBaseValueSourceCell(inventoryRow.SourceCell, out baseRowNumber))
				{
					if (!string.IsNullOrWhiteSpace(inventoryRow.ProposedFieldKey))
					{
						plan.Warnings.Add("FieldInventory row " + inventoryRow.RowNumber.ToString() + " は SourceCell が B列行参照ではないため同期対象外です。 SourceCell=" + (inventoryRow.SourceCell ?? string.Empty));
					}

					continue;
				}

				if (inventoryByBaseRow.ContainsKey(baseRowNumber))
				{
					plan.Errors.Add("CaseList_FieldInventory に同じ SourceCell 行が複数あります。 SourceCell=B" + baseRowNumber.ToString());
					continue;
				}

				inventoryByBaseRow[baseRowNumber] = inventoryRow;
			}

			foreach (BaseHomeFieldKeyRow baseRow in baseRows)
			{
				if (baseRow == null)
				{
					continue;
				}

				plan.CheckedCount++;
				seenBaseRows.Add(baseRow.RowNumber);
				string proposedKey = NormalizeKey(baseRow.FieldKey);
				if (proposedKey.Length == 0)
				{
					plan.Errors.Add("Base HOME A列に空欄があります。 row=" + baseRow.RowNumber.ToString());
					continue;
				}

				if (!string.Equals(baseRow.FieldKey ?? string.Empty, proposedKey, StringComparison.Ordinal))
				{
					plan.Errors.Add("Base HOME A列のキーに前後空白があります。 row=" + baseRow.RowNumber.ToString() + ", key=" + proposedKey);
					continue;
				}

				if (ContainsInvalidCharacter(proposedKey))
				{
					plan.Errors.Add("Base HOME A列のキーに制御文字、改行、タブのいずれかが含まれています。 row=" + baseRow.RowNumber.ToString());
					continue;
				}

				int firstRow;
				if (seenBaseKeys.TryGetValue(proposedKey, out firstRow))
				{
					plan.Errors.Add("Base HOME A列のキーが重複しています。 key=" + proposedKey + ", rows=" + firstRow.ToString() + "," + baseRow.RowNumber.ToString());
					continue;
				}

				seenBaseKeys[proposedKey] = baseRow.RowNumber;

				FieldInventoryRow inventoryRow;
				if (!inventoryByBaseRow.TryGetValue(baseRow.RowNumber, out inventoryRow))
				{
					plan.Errors.Add("Base HOME row " + baseRow.RowNumber.ToString() + " に対応する CaseList_FieldInventory 行が見つかりません。期待 SourceCell=B" + baseRow.RowNumber.ToString());
					continue;
				}

				string oldKey = NormalizeKey(inventoryRow.ProposedFieldKey);
				if (importantKeys.Contains(oldKey)
					&& !string.Equals(oldKey, proposedKey, StringComparison.OrdinalIgnoreCase)
					&& !FieldKeyRenameMap.IsAllowedLegacyToCurrentRename(oldKey, proposedKey))
				{
					plan.Errors.Add("重要キーの変更は自動同期できません。 row=" + baseRow.RowNumber.ToString() + ", old=" + oldKey + ", new=" + proposedKey);
					continue;
				}

				finalKeysByInventoryRow[inventoryRow.RowNumber] = proposedKey;
				if (string.Equals(oldKey, proposedKey, StringComparison.Ordinal))
				{
					plan.UnchangedCount++;
				}
				else
				{
					plan.Updates.Add(new FieldInventoryUpdate
					{
						BaseHomeRowNumber = baseRow.RowNumber,
						FieldInventoryRowNumber = inventoryRow.RowNumber,
						OldFieldKey = oldKey,
						NewFieldKey = proposedKey
					});
				}
			}

			AddFinalDuplicateErrors(plan, finalKeysByInventoryRow);
			AddExtraInventoryWarnings(plan, inventoryRows, seenBaseRows);
			return plan;
		}

		private string ResolveSystemRoot(WorkbookContext context, Excel.Workbook kernelWorkbook)
		{
			string expectedSystemRoot = _pathCompatibilityService.NormalizePath(context == null ? string.Empty : context.SystemRoot);
			if (string.IsNullOrWhiteSpace(expectedSystemRoot))
			{
				throw new InvalidOperationException("SYSTEM_ROOT context was not available for Base HOME field inventory sync.");
			}

			string workbookSystemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(kernelWorkbook, SystemRootPropertyName));
			if (string.IsNullOrWhiteSpace(workbookSystemRoot)
				|| !string.Equals(workbookSystemRoot, expectedSystemRoot, StringComparison.OrdinalIgnoreCase))
			{
				throw new InvalidOperationException("Kernel workbook SYSTEM_ROOT mismatched for Base HOME field inventory sync.");
			}

			return expectedSystemRoot;
		}

		private Excel.Worksheet ResolveBaseHomeWorksheet(Excel.Workbook baseWorkbook)
		{
			Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName(baseWorkbook, BaseHomeSheetCodeName);
			if (worksheet != null)
			{
				return worksheet;
			}

			return _excelInteropService.FindWorksheetByName(baseWorkbook, BaseHomeSheetName);
		}

		private static IReadOnlyList<BaseHomeFieldKeyRow> ReadBaseHomeFieldKeyRows(Excel.Worksheet worksheet)
		{
			var rows = new List<BaseHomeFieldKeyRow>();
			Excel.Range range = null;
			try
			{
				int lastRow = ((dynamic)worksheet.Cells[worksheet.Rows.Count, "A"]).End[Excel.XlDirection.xlUp].Row;
				if (lastRow < 1)
				{
					return rows;
				}

				range = worksheet.Range["A1", "A" + lastRow.ToString()];
				object value = range.Value2;
				if (value is object[,] values)
				{
					for (int rowIndex = 1; rowIndex <= values.GetUpperBound(0); rowIndex++)
					{
						rows.Add(new BaseHomeFieldKeyRow
						{
							RowNumber = rowIndex,
							FieldKey = Convert.ToString(values[rowIndex, 1]) ?? string.Empty
						});
					}
				}
				else
				{
					rows.Add(new BaseHomeFieldKeyRow
					{
						RowNumber = 1,
						FieldKey = Convert.ToString(value) ?? string.Empty
					});
				}

				return rows;
			}
			finally
			{
				ComObjectReleaseService.FinalRelease(range);
			}
		}

		private static IReadOnlyList<FieldInventoryRow> ReadFieldInventoryRows(Excel.Worksheet worksheet, out int proposedFieldKeyColumnIndex)
		{
			var rows = new List<FieldInventoryRow>();
			Excel.Range range = null;
			proposedFieldKeyColumnIndex = 0;
			try
			{
				int lastRow = ((dynamic)worksheet.Cells[worksheet.Rows.Count, 1]).End[Excel.XlDirection.xlUp].Row;
				int lastColumn = ((dynamic)worksheet.Cells[1, worksheet.Columns.Count]).End[Excel.XlDirection.xlToLeft].Column;
				if (lastRow < 2 || lastColumn < 1)
				{
					return rows;
				}

				range = worksheet.Range[(dynamic)worksheet.Cells[1, 1], (dynamic)worksheet.Cells[lastRow, lastColumn]];
				object[,] values = range.Value2 as object[,];
				if (values == null)
				{
					return rows;
				}

				int sourceCellColumnIndex = FindHeaderColumn(values, lastColumn, SourceCellHeaderName);
				proposedFieldKeyColumnIndex = FindHeaderColumn(values, lastColumn, ProposedFieldKeyHeaderName);
				if (sourceCellColumnIndex <= 0 || proposedFieldKeyColumnIndex <= 0)
				{
					throw new InvalidOperationException("CaseList_FieldInventory の SourceCell / ProposedFieldKey 列を特定できません。");
				}

				for (int rowIndex = 2; rowIndex <= values.GetUpperBound(0); rowIndex++)
				{
					string sourceCell = Convert.ToString(values[rowIndex, sourceCellColumnIndex]) ?? string.Empty;
					string proposedFieldKey = Convert.ToString(values[rowIndex, proposedFieldKeyColumnIndex]) ?? string.Empty;
					if (string.IsNullOrWhiteSpace(sourceCell) && string.IsNullOrWhiteSpace(proposedFieldKey))
					{
						continue;
					}

					rows.Add(new FieldInventoryRow
					{
						RowNumber = rowIndex,
						SourceCell = sourceCell,
						ProposedFieldKey = proposedFieldKey
					});
				}

				return rows;
			}
			finally
			{
				ComObjectReleaseService.FinalRelease(range);
			}
		}

		private static int FindHeaderColumn(object[,] values, int lastColumn, string headerName)
		{
			for (int columnIndex = 1; columnIndex <= lastColumn; columnIndex++)
			{
				string text = (Convert.ToString(values[1, columnIndex]) ?? string.Empty).Trim();
				if (string.Equals(text, headerName, StringComparison.OrdinalIgnoreCase))
				{
					return columnIndex;
				}
			}

			return 0;
		}

		private static void ApplyUpdates(Excel.Worksheet worksheet, int proposedFieldKeyColumnIndex, IReadOnlyList<FieldInventoryUpdate> updates)
		{
			if (worksheet == null)
			{
				throw new ArgumentNullException(nameof(worksheet));
			}

			if (proposedFieldKeyColumnIndex <= 0)
			{
				throw new ArgumentOutOfRangeException(nameof(proposedFieldKeyColumnIndex));
			}

			if (updates == null)
			{
				return;
			}

			foreach (FieldInventoryUpdate update in updates)
			{
				if (update == null)
				{
					continue;
				}

				Excel.Range targetCell = null;
				try
				{
					targetCell = worksheet.Cells[update.FieldInventoryRowNumber, proposedFieldKeyColumnIndex] as Excel.Range;
					if (targetCell == null)
					{
						throw new InvalidOperationException("CaseList_FieldInventory の ProposedFieldKey セルを取得できませんでした。");
					}

					targetCell.Value2 = update.NewFieldKey ?? string.Empty;
				}
				finally
				{
					ComObjectReleaseService.FinalRelease(targetCell);
				}
			}
		}

		private static bool TryParseBaseValueSourceCell(string sourceCell, out int baseRowNumber)
		{
			baseRowNumber = 0;
			string text = (sourceCell ?? string.Empty).Trim();
			if (text.Length < 2)
			{
				return false;
			}

			if (text[0] != 'B' && text[0] != 'b')
			{
				return false;
			}

			return int.TryParse(text.Substring(1), out baseRowNumber) && baseRowNumber > 0;
		}

		private static void AddFinalDuplicateErrors(SyncPlan plan, IReadOnlyDictionary<int, string> finalKeysByInventoryRow)
		{
			var rowsByKey = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
			foreach (KeyValuePair<int, string> pair in finalKeysByInventoryRow)
			{
				string key = NormalizeKey(pair.Value);
				if (key.Length == 0)
				{
					continue;
				}

				List<int> rows;
				if (!rowsByKey.TryGetValue(key, out rows))
				{
					rows = new List<int>();
					rowsByKey.Add(key, rows);
				}

				rows.Add(pair.Key);
			}

			foreach (KeyValuePair<string, List<int>> pair in rowsByKey)
			{
				if (pair.Value.Count > 1)
				{
					plan.Errors.Add("同期後の CaseList_FieldInventory.ProposedFieldKey が重複します。 key=" + pair.Key + ", rows=" + string.Join(",", pair.Value.Select(row => row.ToString()).ToArray()));
				}
			}
		}

		private static void AddExtraInventoryWarnings(
			SyncPlan plan,
			IReadOnlyList<FieldInventoryRow> inventoryRows,
			ISet<int> seenBaseRows)
		{
			if (inventoryRows == null)
			{
				return;
			}

			foreach (FieldInventoryRow inventoryRow in inventoryRows)
			{
				if (inventoryRow == null)
				{
					continue;
				}

				int baseRowNumber;
				if (!TryParseBaseValueSourceCell(inventoryRow.SourceCell, out baseRowNumber))
				{
					continue;
				}

				string key = NormalizeKey(inventoryRow.ProposedFieldKey);
				if (key.Length != 0 && (seenBaseRows == null || !seenBaseRows.Contains(baseRowNumber)))
				{
					plan.Warnings.Add("CaseList_FieldInventory row " + inventoryRow.RowNumber.ToString() + " は Base HOME の読取範囲外です。既存行は削除しません。 SourceCell=B" + baseRowNumber.ToString() + ", key=" + key);
				}
			}
		}

		private static bool ContainsInvalidCharacter(string value)
		{
			if (string.IsNullOrEmpty(value))
			{
				return false;
			}

			for (int i = 0; i < value.Length; i++)
			{
				char c = value[i];
				if (c == '\t' || c == '\r' || c == '\n' || char.IsControl(c))
				{
					return true;
				}
			}

			return false;
		}

		private static string NormalizeKey(string value)
		{
			return (value ?? string.Empty).Trim();
		}

		private static SyncResult BuildFailureResult(SyncPlan plan)
		{
			StringBuilder builder = new StringBuilder();
			builder.AppendLine("Base HOME A列から CaseList_FieldInventory への同期を中断しました。");
			builder.AppendLine("Kernel の既存メタ情報は変更していません。");
			builder.AppendLine();
			builder.AppendLine("エラー:");
			foreach (string error in plan.Errors)
			{
				builder.AppendLine("- " + error);
			}

			if (plan.Warnings.Count > 0)
			{
				builder.AppendLine();
				builder.AppendLine("警告:");
				foreach (string warning in plan.Warnings)
				{
					builder.AppendLine("- " + warning);
				}
			}

			return new SyncResult
			{
				Success = false,
				CheckedCount = plan.CheckedCount,
				UpdatedCount = 0,
				UnchangedCount = plan.UnchangedCount,
				WarningCount = plan.Warnings.Count,
				Message = builder.ToString()
			};
		}

		private static SyncResult BuildSuccessResult(SyncPlan plan)
		{
			StringBuilder builder = new StringBuilder();
			builder.AppendLine("Base HOME A列から CaseList_FieldInventory への同期が完了しました。");
			builder.AppendLine("確認行: " + plan.CheckedCount.ToString());
			builder.AppendLine("更新: " + plan.Updates.Count.ToString());
			builder.AppendLine("変更なし: " + plan.UnchangedCount.ToString());
			builder.AppendLine("警告: " + plan.Warnings.Count.ToString());

			if (plan.Updates.Count > 0)
			{
				builder.AppendLine();
				builder.AppendLine("更新内容:");
				foreach (FieldInventoryUpdate update in plan.Updates.Take(20))
				{
					builder.AppendLine("- Base HOME row " + update.BaseHomeRowNumber.ToString() + " / FieldInventory row " + update.FieldInventoryRowNumber.ToString() + ": " + (update.OldFieldKey ?? string.Empty) + " -> " + (update.NewFieldKey ?? string.Empty));
				}

				if (plan.Updates.Count > 20)
				{
					builder.AppendLine("- ...ほか " + (plan.Updates.Count - 20).ToString() + " 件");
				}
			}

			if (plan.Warnings.Count > 0)
			{
				builder.AppendLine();
				builder.AppendLine("警告:");
				foreach (string warning in plan.Warnings)
				{
					builder.AppendLine("- " + warning);
				}
			}

			builder.AppendLine();
			builder.AppendLine("Word 雛形の CC Tag も同じキーへ更新してください。");
			builder.AppendLine("その後、雛形登録・更新を実行して、Kernel 雛形一覧と Base snapshot を更新してください。");

			return new SyncResult
			{
				Success = true,
				CheckedCount = plan.CheckedCount,
				UpdatedCount = plan.Updates.Count,
				UnchangedCount = plan.UnchangedCount,
				WarningCount = plan.Warnings.Count,
				Message = builder.ToString()
			};
		}

		private static SheetProtectionState SaveSheetProtectionState(Excel.Worksheet worksheet)
		{
			SheetProtectionState state = new SheetProtectionState
			{
				IsProtected = worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios
			};
			if (state.IsProtected)
			{
				Excel.Protection protection = worksheet.Protection;
				state.AllowFormattingCells = protection.AllowFormattingCells;
				state.AllowFormattingColumns = protection.AllowFormattingColumns;
				state.AllowFormattingRows = protection.AllowFormattingRows;
				state.AllowInsertingColumns = protection.AllowInsertingColumns;
				state.AllowInsertingRows = protection.AllowInsertingRows;
				state.AllowInsertingHyperlinks = protection.AllowInsertingHyperlinks;
				state.AllowDeletingColumns = protection.AllowDeletingColumns;
				state.AllowDeletingRows = protection.AllowDeletingRows;
				state.AllowSorting = protection.AllowSorting;
				state.AllowFiltering = protection.AllowFiltering;
				state.AllowUsingPivotTables = protection.AllowUsingPivotTables;
			}

			return state;
		}

		private static void RestoreSheetProtectionState(Excel.Worksheet worksheet, SheetProtectionState state)
		{
			if (worksheet == null || state == null || !state.IsProtected)
			{
				return;
			}

			try
			{
				worksheet.Protect(string.Empty, AllowFormattingCells: state.AllowFormattingCells, AllowFormattingColumns: state.AllowFormattingColumns, AllowFormattingRows: state.AllowFormattingRows, AllowInsertingColumns: state.AllowInsertingColumns, AllowInsertingRows: state.AllowInsertingRows, AllowInsertingHyperlinks: state.AllowInsertingHyperlinks, AllowDeletingColumns: state.AllowDeletingColumns, AllowDeletingRows: state.AllowDeletingRows, AllowSorting: state.AllowSorting, AllowFiltering: state.AllowFiltering, AllowUsingPivotTables: state.AllowUsingPivotTables, DrawingObjects: Type.Missing, Contents: Type.Missing, Scenarios: Type.Missing, UserInterfaceOnly: true);
				worksheet.EnableSelection = Excel.XlEnableSelection.xlUnlockedCells;
			}
			catch
			{
			}
		}
	}
}
