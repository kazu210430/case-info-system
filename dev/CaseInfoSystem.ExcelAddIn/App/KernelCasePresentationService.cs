using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelCasePresentationService
	{
		private static readonly IntPtr HwndTopMost = new IntPtr (-1);

		private static readonly IntPtr HwndNoTopMost = new IntPtr (-2);

		private const int SwRestore = 9;

		private const int PromoteRetryCount = 4;

		private const int PromoteRetryIntervalMs = 150;

		private const uint SwpNoMove = 2u;

		private const uint SwpNoSize = 1u;

		private const uint SwpShowWindow = 64u;

		private const string HomeSheetCodeName = "shHOME";

		private const string HomeSheetName = "ホーム";

		private const string InitialCursorFieldKey = "顧客_よみ";

		private readonly Application _application;

		private readonly CaseWorkbookOpenStrategy _caseWorkbookOpenStrategy;

		private readonly ExcelInteropService _excelInteropService;

		private readonly ExcelWindowRecoveryService _excelWindowRecoveryService;

		private readonly KernelWorkbookResolverService _kernelWorkbookResolverService;

		private readonly CaseListFieldDefinitionRepository _caseListFieldDefinitionRepository;

		private readonly FolderWindowService _folderWindowService;

		private readonly TransientPaneSuppressionService _transientPaneSuppressionService;

		private readonly Logger _logger;

		[DllImport ("user32.dll")]
		private static extern bool ShowWindow (IntPtr hWnd, int nCmdShow);

		[DllImport ("user32.dll")]
		private static extern bool SetForegroundWindow (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern bool SetWindowPos (IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

		internal KernelCasePresentationService (Application application, CaseWorkbookOpenStrategy caseWorkbookOpenStrategy, ExcelInteropService excelInteropService, ExcelWindowRecoveryService excelWindowRecoveryService, KernelWorkbookResolverService kernelWorkbookResolverService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository, FolderWindowService folderWindowService, TransientPaneSuppressionService transientPaneSuppressionService, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_caseWorkbookOpenStrategy = caseWorkbookOpenStrategy ?? throw new ArgumentNullException ("caseWorkbookOpenStrategy");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_excelWindowRecoveryService = excelWindowRecoveryService ?? throw new ArgumentNullException ("excelWindowRecoveryService");
			_kernelWorkbookResolverService = kernelWorkbookResolverService ?? throw new ArgumentNullException ("kernelWorkbookResolverService");
			_caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException ("caseListFieldDefinitionRepository");
			_folderWindowService = folderWindowService ?? throw new ArgumentNullException ("folderWindowService");
			_transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException ("transientPaneSuppressionService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void OpenCaseFolder (string caseFolderPath, string reason)
		{
			if (string.IsNullOrWhiteSpace (caseFolderPath)) {
				return;
			}
			try {
				_folderWindowService.OpenFolder (caseFolderPath, reason ?? "KernelCasePresentationService.OpenCaseFolder");
			} catch (Exception exception) {
				_logger.Error ("OpenCaseFolder failed.", exception);
			}
		}

		internal KernelCaseCreationResult OpenCreatedCase (KernelCaseCreationResult result)
		{
			if (result == null) {
				throw new ArgumentNullException ("result");
			}
			if (!result.Success) {
				return result;
			}
			if (string.IsNullOrWhiteSpace (result.CaseWorkbookPath)) {
				throw new InvalidOperationException ("CASE workbook path could not be resolved.");
			}
			Stopwatch stopwatch = Stopwatch.StartNew ();
			Workbook workbook = null;
			try {
				_caseWorkbookOpenStrategy.RegisterKnownCasePath (result.CaseWorkbookPath);
				_transientPaneSuppressionService.SuppressPath (result.CaseWorkbookPath, "KernelCasePresentationService.OpenCreatedCase");
				workbook = _caseWorkbookOpenStrategy.OpenVisibleWorkbook (result.CaseWorkbookPath);
				_logger.Info ("Kernel prompt CASE workbook opened. path=" + result.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				if (workbook == null) {
					throw new InvalidOperationException ("CASE workbook could not be opened.");
				}
				result.CreatedWorkbook = workbook;
				ShowCreatedCase (workbook);
				_logger.Info ("Kernel prompt CASE presentation completed. path=" + result.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return result;
			} catch {
				_transientPaneSuppressionService.ReleasePath (result.CaseWorkbookPath, "KernelCasePresentationService.OpenCreatedCaseFailure");
				throw;
			}
		}

		private void ShowCreatedCase (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			try {
				Stopwatch stopwatch = Stopwatch.StartNew ();
				_excelWindowRecoveryService.TryRecoverWorkbookWindow (workbook, "KernelCasePresentationService.ShowCreatedCase", bringToFront: true);
				_logger.Info ("ShowCreatedCase workbook activated. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				Globals.ThisAddIn.RefreshActiveTaskPane ("KernelCasePresentationService.ShowCreatedCase");
				_logger.Info ("ShowCreatedCase task pane refreshed. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				PromoteWorkbookWindow (workbook);
				_logger.Info ("ShowCreatedCase window promoted. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_transientPaneSuppressionService.ReleaseWorkbook (workbook, "KernelCasePresentationService.ShowCreatedCase");
				Globals.ThisAddIn.RefreshActiveTaskPane ("KernelCasePresentationService.ShowCreatedCase.PostRelease");
				_logger.Info ("ShowCreatedCase task pane post-release refreshed. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				Globals.ThisAddIn.SuppressUpcomingCasePaneActivationRefresh (_excelInteropService.GetWorkbookFullName (workbook), "KernelCasePresentationService.ShowCreatedCase.PostRelease");
				_logger.Info ("ShowCreatedCase post-release activation suppression prepared. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				MoveInitialCursorToHomeCell (workbook);
				_logger.Info ("ShowCreatedCase cursor positioned. elapsedMs=" + stopwatch.ElapsedMilliseconds);
			} catch (Exception exception) {
				_logger.Error ("ShowCreatedCase failed.", exception);
			}
		}

		private void MoveInitialCursorToHomeCell (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (workbook, "shHOME");
			if (worksheet == null) {
				try {
					worksheet = workbook.Worksheets ["ホーム"] as Worksheet;
				} catch {
					worksheet = workbook.Worksheets [1] as Worksheet;
				}
			}
			if (worksheet == null) {
				return;
			}
			worksheet.Activate ();
			Range range = ResolveInitialCursorRange (workbook, worksheet);
			if (range == null) {
				return;
			}
			try {
				range.Select ();
			} finally {
				Marshal.FinalReleaseComObject (range);
			}
		}

		private Range ResolveInitialCursorRange (Workbook caseWorkbook, Worksheet homeWorksheet)
		{
			if (caseWorkbook == null || homeWorksheet == null) {
				return null;
			}
			bool openedNow;
			Workbook workbook = _kernelWorkbookResolverService.ResolveOrOpenReadOnly (caseWorkbook, out openedNow);
			if (workbook == null) {
				return null;
			}
			try {
				IReadOnlyDictionary<string, CaseListFieldDefinition> readOnlyDictionary = _caseListFieldDefinitionRepository.LoadDefinitions (workbook);
				if (readOnlyDictionary == null) {
					return null;
				}
				readOnlyDictionary.TryGetValue ("顧客_よみ", out var value);
				return _excelInteropService.ResolveFieldRange (caseWorkbook, homeWorksheet, value);
			} finally {
				if (openedNow && workbook != null) {
					try {
						workbook.Close (false, Type.Missing, Type.Missing);
					} catch (Exception exception) {
						_logger.Error ("ResolveInitialCursorRange temporary kernel close failed.", exception);
					}
				}
			}
		}

		private void PromoteWorkbookWindow (Workbook workbook)
		{
			Window firstVisibleWindow = _excelInteropService.GetFirstVisibleWindow (workbook);
			if (firstVisibleWindow == null) {
				return;
			}
			try {
				IntPtr hwnd = new IntPtr (firstVisibleWindow.Hwnd);
				IntPtr hwnd2 = new IntPtr (_application.Hwnd);
				for (int i = 0; i < 4; i++) {
					PromoteWindow (hwnd2);
					PromoteWindow (hwnd);
					if (i < 3) {
						Thread.Sleep (150);
					}
				}
				_logger.Info ("Created CASE workbook window promoted. hwnd=" + firstVisibleWindow.Hwnd);
			} catch (Exception exception) {
				_logger.Error ("PromoteWorkbookWindow failed.", exception);
			}
		}

		private static void PromoteWindow (IntPtr hwnd)
		{
			if (!(hwnd == IntPtr.Zero)) {
				ShowWindow (hwnd, 9);
				SetWindowPos (hwnd, HwndTopMost, 0, 0, 0, 0, 67u);
				SetWindowPos (hwnd, HwndNoTopMost, 0, 0, 0, 0, 67u);
				SetForegroundWindow (hwnd);
			}
		}
	}
}
