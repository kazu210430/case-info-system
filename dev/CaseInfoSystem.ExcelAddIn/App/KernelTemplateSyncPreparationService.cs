using System;
using System.Collections.Generic;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelTemplateSyncPreparationService
	{
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
			private readonly Worksheet _worksheet;

			private readonly SheetProtectionState _state;

			private bool _disposed;

			internal TemporarySheetProtectionRestoreScope (Worksheet worksheet)
			{
				_worksheet = worksheet ?? throw new ArgumentNullException ("worksheet");
				_state = SaveSheetProtectionState (worksheet);
				if (_state.IsProtected) {
					worksheet.Unprotect (string.Empty);
				}
			}

			public void Dispose ()
			{
				if (_disposed) {
					return;
				}
				_disposed = true;
				RestoreSheetProtectionState (_worksheet, _state);
			}
		}

		internal sealed class PreparedKernelTemplateSyncScope : IDisposable
		{
			private readonly IDisposable _protectionRestoreScope;

			private bool _disposed;

			internal Worksheet MasterSheet { get; }

			internal string SystemRoot { get; }

			internal KernelTemplateSyncPreflightResult PreflightResult { get; }

			internal PreparedKernelTemplateSyncScope (Worksheet masterSheet, string systemRoot, KernelTemplateSyncPreflightResult preflightResult, IDisposable protectionRestoreScope)
			{
				MasterSheet = masterSheet ?? throw new ArgumentNullException ("masterSheet");
				SystemRoot = systemRoot ?? string.Empty;
				PreflightResult = preflightResult ?? throw new ArgumentNullException ("preflightResult");
				_protectionRestoreScope = protectionRestoreScope ?? throw new ArgumentNullException ("protectionRestoreScope");
			}

			public void Dispose ()
			{
				if (_disposed) {
					return;
				}
				_disposed = true;
				_protectionRestoreScope.Dispose ();
			}
		}

		private const string MasterSheetCodeName = "shMasterList";

		private const string MasterSheetName = "雛形一覧";

		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly CaseListFieldDefinitionRepository _caseListFieldDefinitionRepository;

		private readonly KernelTemplateSyncPreflightService _kernelTemplateSyncPreflightService;

		internal KernelTemplateSyncPreparationService (ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository, KernelTemplateSyncPreflightService kernelTemplateSyncPreflightService)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException ("caseListFieldDefinitionRepository");
			_kernelTemplateSyncPreflightService = kernelTemplateSyncPreflightService ?? throw new ArgumentNullException ("kernelTemplateSyncPreflightService");
		}

		internal PreparedKernelTemplateSyncScope Prepare (Workbook kernelWorkbook)
		{
			if (kernelWorkbook == null) {
				throw new ArgumentNullException ("kernelWorkbook");
			}
			Worksheet masterSheet = GetMasterListSheet (kernelWorkbook);
			TemporarySheetProtectionRestoreScope protectionRestoreScope = new TemporarySheetProtectionRestoreScope (masterSheet);
			try {
				ValidateMasterListSheet (masterSheet);
				string systemRoot = ResolveSystemRoot (kernelWorkbook);
				KernelTemplateSyncPreflightResult preflightResult = _kernelTemplateSyncPreflightService.Run (new KernelTemplateSyncPreflightRequest (systemRoot, LoadDefinedTemplateTags (kernelWorkbook)));
				return new PreparedKernelTemplateSyncScope (masterSheet, systemRoot, preflightResult, protectionRestoreScope);
			} catch {
				protectionRestoreScope.Dispose ();
				throw;
			}
		}

		private IReadOnlyCollection<string> LoadDefinedTemplateTags (Workbook kernelWorkbook)
		{
			IReadOnlyDictionary<string, CaseListFieldDefinition> readOnlyDictionary = _caseListFieldDefinitionRepository.LoadDefinitions (kernelWorkbook);
			if (readOnlyDictionary == null || readOnlyDictionary.Count == 0) {
				return Array.Empty<string> ();
			}
			return readOnlyDictionary.Keys
				.Where (key => !string.IsNullOrWhiteSpace (key))
				.Select (key => key.Trim ())
				.ToArray ();
		}

		private string ResolveSystemRoot (Workbook kernelWorkbook)
		{
			string text = _pathCompatibilityService.NormalizePath (_excelInteropService.TryGetDocumentProperty (kernelWorkbook, "SYSTEM_ROOT"));
			if (!string.IsNullOrWhiteSpace (text)) {
				return text;
			}
			text = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookPath (kernelWorkbook));
			if (string.IsNullOrWhiteSpace (text)) {
				throw new InvalidOperationException ("SYSTEM_ROOT を解決できませんでした。");
			}
			return text;
		}

		private Worksheet GetMasterListSheet (Workbook kernelWorkbook)
		{
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (kernelWorkbook, MasterSheetCodeName);
			if (worksheet != null) {
				return worksheet;
			}
			try {
				worksheet = kernelWorkbook.Worksheets [MasterSheetName] as Worksheet;
			} catch {
				worksheet = null;
			}
			if (worksheet == null) {
				throw new InvalidOperationException ("雛形登録シートが見つかりません。");
			}
			return worksheet;
		}

		private static void ValidateMasterListSheet (Worksheet worksheet)
		{
			if (worksheet == null) {
				throw new ArgumentNullException ("worksheet");
			}
			if (worksheet.ProtectContents) {
				throw new InvalidOperationException ("雛形登録シートが保護されています。保護解除してから実行してください。");
			}
		}

		private static SheetProtectionState SaveSheetProtectionState (Worksheet worksheet)
		{
			SheetProtectionState sheetProtectionState = new SheetProtectionState {
				IsProtected = (worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios)
			};
			if (sheetProtectionState.IsProtected) {
				Protection protection = worksheet.Protection;
				sheetProtectionState.AllowFormattingCells = protection.AllowFormattingCells;
				sheetProtectionState.AllowFormattingColumns = protection.AllowFormattingColumns;
				sheetProtectionState.AllowFormattingRows = protection.AllowFormattingRows;
				sheetProtectionState.AllowInsertingColumns = protection.AllowInsertingColumns;
				sheetProtectionState.AllowInsertingRows = protection.AllowInsertingRows;
				sheetProtectionState.AllowInsertingHyperlinks = protection.AllowInsertingHyperlinks;
				sheetProtectionState.AllowDeletingColumns = protection.AllowDeletingColumns;
				sheetProtectionState.AllowDeletingRows = protection.AllowDeletingRows;
				sheetProtectionState.AllowSorting = protection.AllowSorting;
				sheetProtectionState.AllowFiltering = protection.AllowFiltering;
				sheetProtectionState.AllowUsingPivotTables = protection.AllowUsingPivotTables;
			}
			return sheetProtectionState;
		}

		private static void RestoreSheetProtectionState (Worksheet worksheet, SheetProtectionState state)
		{
			if (worksheet == null || state == null || !state.IsProtected) {
				return;
			}
			try {
				worksheet.Protect (string.Empty, AllowFormattingCells: state.AllowFormattingCells, AllowFormattingColumns: state.AllowFormattingColumns, AllowFormattingRows: state.AllowFormattingRows, AllowInsertingColumns: state.AllowInsertingColumns, AllowInsertingRows: state.AllowInsertingRows, AllowInsertingHyperlinks: state.AllowInsertingHyperlinks, AllowDeletingColumns: state.AllowDeletingColumns, AllowDeletingRows: state.AllowDeletingRows, AllowSorting: state.AllowSorting, AllowFiltering: state.AllowFiltering, AllowUsingPivotTables: state.AllowUsingPivotTables, DrawingObjects: Type.Missing, Contents: Type.Missing, Scenarios: Type.Missing, UserInterfaceOnly: true);
				worksheet.EnableSelection = XlEnableSelection.xlUnlockedCells;
			} catch {
			}
		}
	}
}
