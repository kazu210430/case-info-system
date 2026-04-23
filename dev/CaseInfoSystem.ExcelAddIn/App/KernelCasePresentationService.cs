using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
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

		private readonly CreatedCasePresentationWaitService _createdCasePresentationWaitService;

		private readonly TransientPaneSuppressionService _transientPaneSuppressionService;

		private readonly Logger _logger;

		[DllImport ("user32.dll")]
		private static extern bool ShowWindow (IntPtr hWnd, int nCmdShow);

		[DllImport ("user32.dll")]
		private static extern bool SetForegroundWindow (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern bool SetWindowPos (IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

		internal KernelCasePresentationService (Application application, CaseWorkbookOpenStrategy caseWorkbookOpenStrategy, ExcelInteropService excelInteropService, ExcelWindowRecoveryService excelWindowRecoveryService, KernelWorkbookResolverService kernelWorkbookResolverService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository, FolderWindowService folderWindowService, CreatedCasePresentationWaitService createdCasePresentationWaitService, TransientPaneSuppressionService transientPaneSuppressionService, Logger logger)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_caseWorkbookOpenStrategy = caseWorkbookOpenStrategy ?? throw new ArgumentNullException ("caseWorkbookOpenStrategy");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_excelWindowRecoveryService = excelWindowRecoveryService ?? throw new ArgumentNullException ("excelWindowRecoveryService");
			_kernelWorkbookResolverService = kernelWorkbookResolverService ?? throw new ArgumentNullException ("kernelWorkbookResolverService");
			_caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException ("caseListFieldDefinitionRepository");
			_folderWindowService = folderWindowService ?? throw new ArgumentNullException ("folderWindowService");
			_createdCasePresentationWaitService = createdCasePresentationWaitService ?? throw new ArgumentNullException ("createdCasePresentationWaitService");
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
			CreatedCasePresentationWaitService.WaitSession waitSession = _createdCasePresentationWaitService.ShowWaiting (stopwatch);
			try {
				TryOpenCaseFolderBeforeShowingCase (result.CaseFolderPath, result.CaseWorkbookPath);
				_caseWorkbookOpenStrategy.RegisterKnownCasePath (result.CaseWorkbookPath);
				_transientPaneSuppressionService.SuppressPath (result.CaseWorkbookPath, "KernelCasePresentationService.OpenCreatedCase");
				workbook = OpenCreatedCaseWorkbook (result);
				_logger.Info ("Kernel prompt CASE workbook opened. path=" + result.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				if (workbook == null) {
					throw new InvalidOperationException ("CASE workbook could not be opened.");
				}
				result.CreatedWorkbook = workbook;
				ShowCreatedCase (workbook);
				waitSession.CloseForSuccessfulPresentation ();
				PromoteWorkbookWindowOnce (workbook, "KernelCasePresentationService.OpenCreatedCase.AfterWaitUiClose");
				_logger.Info ("Kernel prompt CASE presentation completed. path=" + result.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return result;
			} catch {
				waitSession.CloseAndRestoreOwner ();
				_transientPaneSuppressionService.ReleasePath (result.CaseWorkbookPath, "KernelCasePresentationService.OpenCreatedCaseFailure");
				throw;
			} finally {
				waitSession.Dispose ();
			}
		}

		private Workbook OpenCreatedCaseWorkbook (KernelCaseCreationResult result)
		{
			if (ShouldUseHiddenOpenForCreatedCase (result.Mode)) {
				_logger.Info ("Kernel prompt CASE workbook hidden open selected. mode=" + result.Mode.ToString () + ", path=" + result.CaseWorkbookPath);
				return _caseWorkbookOpenStrategy.OpenHiddenForCaseDisplay (result.CaseWorkbookPath);
			}
			_logger.Info ("Kernel prompt CASE workbook visible open selected. mode=" + result.Mode.ToString () + ", path=" + result.CaseWorkbookPath);
			return _caseWorkbookOpenStrategy.OpenVisibleWorkbook (result.CaseWorkbookPath);
		}

		private static bool ShouldUseHiddenOpenForCreatedCase (KernelCaseCreationMode mode)
		{
			return mode == KernelCaseCreationMode.NewCaseDefault || mode == KernelCaseCreationMode.CreateCaseSingle;
		}

		private void TryOpenCaseFolderBeforeShowingCase (string caseFolderPath, string caseWorkbookPath)
		{
			if (string.IsNullOrWhiteSpace (caseFolderPath)) {
				return;
			}
			try {
				IntPtr intPtr = _folderWindowService.OpenFolderAndWait (caseFolderPath, "KernelCasePresentationService.OpenCreatedCase.PreOpen");
				_logger.Info ("CASE folder pre-open completed. workbookPath=" + (caseWorkbookPath ?? string.Empty) + ", folderPath=" + caseFolderPath + ", explorerWindowFound=" + (intPtr != IntPtr.Zero));
			} catch (Exception exception) {
				_logger.Warn ("CASE folder pre-open failed but CASE opening continues. workbookPath=" + (caseWorkbookPath ?? string.Empty) + ", folderPath=" + caseFolderPath + ", message=" + exception.Message);
			}
		}

		private void ShowCreatedCase (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			try {
				Stopwatch stopwatch = Stopwatch.StartNew ();
				_excelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing (workbook, "KernelCasePresentationService.ShowCreatedCase", bringToFront: false);
				_logger.Info ("[KernelFlickerTrace] source=KernelCasePresentationService action=display-stability-point phase=InitialRecoveryCompleted, workbook=" + _excelInteropService.GetWorkbookFullName (workbook) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_logger.Info ("ShowCreatedCase workbook activated. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				ExecuteDeferredPresentationEnhancements (workbook, stopwatch);
			} catch (Exception exception) {
				_logger.Error ("ShowCreatedCase failed.", exception);
			}
		}

		private void ExecuteDeferredPresentationEnhancements (Workbook workbook, Stopwatch stopwatch)
		{
			if (workbook == null) {
				return;
			}
			bool flag = false;
			try {
				_logger.Info ("ShowCreatedCase deferred presentation started. elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
				_transientPaneSuppressionService.ReleaseWorkbook (workbook, "KernelCasePresentationService.ShowCreatedCase");
				flag = true;
				Globals.ThisAddIn.SuppressUpcomingCasePaneActivationRefresh (_excelInteropService.GetWorkbookFullName (workbook), "KernelCasePresentationService.ShowCreatedCase.PostRelease");
				_logger.Info ("ShowCreatedCase post-release activation suppression prepared. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				Globals.ThisAddIn.ShowWorkbookTaskPaneWhenReady (workbook, "KernelCasePresentationService.ShowCreatedCase.PostRelease");
				_logger.Info ("[KernelFlickerTrace] source=KernelCasePresentationService action=display-stability-point phase=ReadyShowRequested, workbook=" + _excelInteropService.GetWorkbookFullName (workbook) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_logger.Info ("ShowCreatedCase task pane ready-show requested. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				try {
					MoveInitialCursorToHomeCell (workbook);
					_logger.Info ("ShowCreatedCase cursor positioned. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				} catch (Exception exception2) {
					_logger.Warn ("ShowCreatedCase cursor positioning skipped after deferred presentation. message=" + exception2.Message + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				}
				_logger.Info ("[KernelFlickerTrace] source=KernelCasePresentationService action=display-stability-point phase=DeferredPresentationCompleted, workbook=" + _excelInteropService.GetWorkbookFullName (workbook) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_logger.Info ("ShowCreatedCase deferred presentation completed. elapsedMs=" + stopwatch.ElapsedMilliseconds);
			} catch (Exception exception) {
				_logger.Error ("ShowCreatedCase deferred presentation failed.", exception);
			} finally {
				if (!flag) {
					try {
						_transientPaneSuppressionService.ReleaseWorkbook (workbook, "KernelCasePresentationService.ShowCreatedCase.DeferredCleanup");
					} catch (Exception exception2) {
						_logger.Error ("ShowCreatedCase deferred cleanup failed.", exception2);
					}
				}
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

		private void PromoteWorkbookWindowOnce (Workbook workbook, string reason)
		{
			if (workbook == null) {
				return;
			}
			Window firstVisibleWindow = _excelInteropService.GetFirstVisibleWindow (workbook);
			if (firstVisibleWindow == null) {
				_logger.Warn ("Created CASE workbook promotion skipped because visible workbook window could not be resolved. reason=" + (reason ?? string.Empty));
				return;
			}
			try {
				IntPtr hwnd = new IntPtr (firstVisibleWindow.Hwnd);
				IntPtr hwnd2 = new IntPtr (_application.Hwnd);
				PromoteWindow (hwnd2);
				PromoteWindow (hwnd);
				_logger.Info ("Created CASE workbook window promoted once. reason=" + (reason ?? string.Empty) + ", workbookHwnd=" + firstVisibleWindow.Hwnd);
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
