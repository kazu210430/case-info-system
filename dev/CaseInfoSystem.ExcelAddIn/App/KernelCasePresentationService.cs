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

		private readonly ICasePaneHostBridge _casePaneHostBridge;

		private readonly WorkbookWindowVisibilityService _workbookWindowVisibilityService;

		private readonly Logger _logger;

		[DllImport ("user32.dll")]
		private static extern bool ShowWindow (IntPtr hWnd, int nCmdShow);

		[DllImport ("user32.dll")]
		private static extern bool SetForegroundWindow (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern bool SetWindowPos (IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

		internal KernelCasePresentationService (Application application, CaseWorkbookOpenStrategy caseWorkbookOpenStrategy, ExcelInteropService excelInteropService, ExcelWindowRecoveryService excelWindowRecoveryService, KernelWorkbookResolverService kernelWorkbookResolverService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository, FolderWindowService folderWindowService, CreatedCasePresentationWaitService createdCasePresentationWaitService, TransientPaneSuppressionService transientPaneSuppressionService, ICasePaneHostBridge casePaneHostBridge, WorkbookWindowVisibilityService workbookWindowVisibilityService, Logger logger)
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
			_casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException ("casePaneHostBridge");
			_workbookWindowVisibilityService = workbookWindowVisibilityService ?? throw new ArgumentNullException ("workbookWindowVisibilityService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void OpenCaseFolderAndWait (string caseFolderPath, string reason)
		{
			if (string.IsNullOrWhiteSpace (caseFolderPath)) {
				return;
			}
			try {
				IntPtr intPtr = _folderWindowService.OpenFolderAndWait (caseFolderPath, reason ?? "KernelCasePresentationService.OpenCaseFolderAndWait");
				_logger.Info ("CASE folder open-and-wait completed. folderPath=" + caseFolderPath + ", explorerWindowFound=" + (intPtr != IntPtr.Zero));
			} catch (Exception exception) {
				_logger.Error ("OpenCaseFolderAndWait failed.", exception);
			}
		}

		internal KernelCaseCreationResult OpenCreatedCase (KernelCaseCreationResult result, CreatedCasePresentationWaitService.WaitSession existingWaitSession = null)
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
			CreatedCasePresentationWaitService.WaitSession waitSession = existingWaitSession ?? _createdCasePresentationWaitService.ShowWaiting (stopwatch);
			try {
				if (existingWaitSession != null) {
					_logger.Info ("Created CASE presentation wait UI reused. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				}
				if (result.Mode == KernelCaseCreationMode.NewCaseDefault) {
					NewCaseDefaultTimingLogHelper.BeginPresentation (result.CaseWorkbookPath);
				}
				_caseWorkbookOpenStrategy.RegisterKnownCasePath (result.CaseWorkbookPath);
				_transientPaneSuppressionService.SuppressPath (result.CaseWorkbookPath, "KernelCasePresentationService.OpenCreatedCase");
				waitSession.UpdateStage (CreatedCasePresentationWaitService.PreparingOpenStageTitle);
				_logger.Info ("Created CASE workbook open started. path=" + result.CaseWorkbookPath + NewCaseVisibilityObservation.FormatCorrelationFields (_excelInteropService, null, result.CaseWorkbookPath));
				NewCaseVisibilityObservation.Log (_logger, _excelInteropService, null, null, null, "case-workbook-open-started", "KernelCasePresentationService.OpenCreatedCase", result.CaseWorkbookPath);
				workbook = OpenCreatedCaseWorkbook (result);
				_logger.Info ("Kernel prompt CASE workbook opened. path=" + result.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds + NewCaseVisibilityObservation.FormatCorrelationFields (_excelInteropService, workbook, result.CaseWorkbookPath));
				if (workbook == null) {
					throw new InvalidOperationException ("CASE workbook could not be opened.");
				}
				string workbookFullName = _excelInteropService.GetWorkbookFullName (workbook);
				NewCaseVisibilityObservation.AttachAlias (result.CaseWorkbookPath, workbookFullName);
				NewCaseVisibilityObservation.Log (_logger, _excelInteropService, null, workbook, null, "case-workbook-open-completed", "KernelCasePresentationService.OpenCreatedCase", result.CaseWorkbookPath);
				if (result.Mode == KernelCaseCreationMode.NewCaseDefault) {
					NewCaseDefaultTimingLogHelper.AttachWorkbookAlias (result.CaseWorkbookPath, workbookFullName);
				}
				result.CreatedWorkbook = workbook;
				CasePresentationResult presentationResult = ShowCreatedCase (workbook, waitSession);
				ApplyPresentationOutcome (result, presentationResult);
				if (result.Mode == KernelCaseCreationMode.NewCaseDefault) {
					NewCaseDefaultTimingLogHelper.StartWaitUiCloseToFinalForegroundStable (result.CaseWorkbookPath);
				}
				waitSession.CloseForSuccessfulPresentation ();
				if (result.Mode != KernelCaseCreationMode.NewCaseDefault) {
					PromoteWorkbookWindowOnce (workbook, "KernelCasePresentationService.OpenCreatedCase.AfterWaitUiClose");
				} else {
					NewCaseDefaultTimingLogHelper.LogWaitUiCloseToFinalForegroundStableIfPending (_logger, result.CaseWorkbookPath, "presentationOutcome=" + result.PresentationOutcome.ToString ());
				}
				_logger.Info ("Kernel prompt CASE presentation completed. path=" + result.CaseWorkbookPath + ", presentationOutcome=" + result.PresentationOutcome.ToString () + ", presentationOutcomeReason=" + SanitizeOutcomeReason (result.PresentationOutcomeReason) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return result;
			} catch (Exception exception) {
				ApplyPresentationOutcome (result, CasePresentationResult.Failed ("OpenCreatedCaseException:" + exception.GetType ().Name));
				_logger.Error ("Kernel prompt CASE presentation failed. path=" + result.CaseWorkbookPath + ", presentationOutcome=" + result.PresentationOutcome.ToString () + ", presentationOutcomeReason=" + SanitizeOutcomeReason (result.PresentationOutcomeReason), exception);
				NewCaseVisibilityObservation.Complete (result.CaseWorkbookPath);
				if (result.Mode == KernelCaseCreationMode.NewCaseDefault) {
					NewCaseDefaultTimingLogHelper.Clear (result.CaseWorkbookPath);
				}
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

		private CasePresentationResult ShowCreatedCase (Workbook workbook, CreatedCasePresentationWaitService.WaitSession waitSession)
		{
			if (workbook == null) {
				return CasePresentationResult.Failed ("WorkbookMissing");
			}
			CasePresentationResult presentationResult = CasePresentationResult.Completed ();
			try {
				Stopwatch stopwatch = Stopwatch.StartNew ();
				string workbookFullName = _excelInteropService.GetWorkbookFullName (workbook);
				Stopwatch stopwatch2 = Stopwatch.StartNew ();
				WorkbookWindowVisibilityEnsureResult initialVisibilityResult = EnsureWorkbookWindowVisibleBeforeInitialRecovery (workbook, stopwatch);
				MarkVisibilityDegradationIfNeeded (presentationResult, "InitialVisibility", initialVisibilityResult);
				NewCaseDefaultTimingLogHelper.LogDetail (_logger, workbookFullName, "hiddenOpenToWindowVisible", "ensureWorkbookWindowVisibleBeforeInitialRecovery", stopwatch2.ElapsedMilliseconds);
				stopwatch2 = Stopwatch.StartNew ();
				bool initialRecoveryCompleted = _excelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing (workbook, "KernelCasePresentationService.ShowCreatedCase", bringToFront: false);
				if (!initialRecoveryCompleted) {
					presentationResult.MarkDegraded ("InitialRecoveryNotCompleted");
				}
				NewCaseDefaultTimingLogHelper.LogDetail (_logger, workbookFullName, "hiddenOpenToWindowVisible", "tryRecoverWorkbookWindowWithoutShowing", stopwatch2.ElapsedMilliseconds);
				_logger.Info ("[KernelFlickerTrace] source=KernelCasePresentationService action=display-stability-point phase=InitialRecoveryCompleted, workbook=" + workbookFullName + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				NewCaseVisibilityObservation.Log (_logger, _excelInteropService, null, workbook, null, "initial-recovery-completed", "KernelCasePresentationService.ShowCreatedCase", workbookFullName);
				_logger.Info ("ShowCreatedCase workbook activated. recovered=" + initialRecoveryCompleted + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				presentationResult.Merge (ExecuteDeferredPresentationEnhancements (workbook, stopwatch, waitSession));
			} catch (Exception exception) {
				_logger.Error ("ShowCreatedCase failed.", exception);
				presentationResult.MarkFailed ("ShowCreatedCaseException:" + exception.GetType ().Name);
			}
			return presentationResult;
		}

		private WorkbookWindowVisibilityEnsureResult EnsureWorkbookWindowVisibleBeforeInitialRecovery (Workbook workbook, Stopwatch stopwatch)
		{
			if (workbook == null) {
				return null;
			}
			WorkbookWindowVisibilityEnsureResult result = _workbookWindowVisibilityService.EnsureVisible (workbook, "KernelCasePresentationService.EnsureWorkbookWindowVisibleBeforeInitialRecovery");
			if (result.Outcome == WorkbookWindowVisibilityEnsureOutcome.MadeVisible) {
				_logger.Info ("ShowCreatedCase workbook window primed before shared application visibility recovery. workbook=" + result.WorkbookFullName + ", windowHwnd=" + result.WindowHwnd + ", elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
			}
			return result;
		}

		private CasePresentationResult ExecuteDeferredPresentationEnhancements (Workbook workbook, Stopwatch stopwatch, CreatedCasePresentationWaitService.WaitSession waitSession)
		{
			if (workbook == null) {
				return CasePresentationResult.Failed ("WorkbookMissing");
			}
			CasePresentationResult presentationResult = CasePresentationResult.Completed ();
			bool flag = false;
				try {
					_logger.Info ("ShowCreatedCase deferred presentation started. elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
					_transientPaneSuppressionService.ReleaseWorkbook (workbook, "KernelCasePresentationService.ShowCreatedCase");
					flag = true;
					Stopwatch stopwatch2 = Stopwatch.StartNew ();
					waitSession?.UpdateStage (CreatedCasePresentationWaitService.ShowingScreenStageTitle);
					WorkbookWindowVisibilityEnsureResult readyShowVisibilityResult = EnsureWorkbookWindowVisibleBeforeReadyShow (workbook, stopwatch);
					MarkVisibilityDegradationIfNeeded (presentationResult, "ReadyShowVisibility", readyShowVisibilityResult);
					NewCaseDefaultTimingLogHelper.LogDetail (_logger, _excelInteropService.GetWorkbookFullName (workbook), "hiddenOpenToWindowVisible", "ensureWorkbookWindowVisibleBeforeReadyShow", stopwatch2.ElapsedMilliseconds);
					_casePaneHostBridge.SuppressUpcomingCasePaneActivationRefresh (_excelInteropService.GetWorkbookFullName (workbook), ControlFlowReasons.CreatedCasePostRelease);
				_logger.Info ("ShowCreatedCase post-release activation suppression prepared. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				NewCaseVisibilityObservation.Log (_logger, _excelInteropService, null, workbook, null, "post-release-suppression-prepared", "KernelCasePresentationService.ExecuteDeferredPresentationEnhancements", _excelInteropService.GetWorkbookFullName (workbook));
				NewCaseDefaultTimingLogHelper.StartTaskPaneReadyWait (_excelInteropService.GetWorkbookFullName (workbook));
				_casePaneHostBridge.ShowWorkbookTaskPaneWhenReady (workbook, ControlFlowReasons.CreatedCasePostRelease);
				_logger.Info ("[KernelFlickerTrace] source=KernelCasePresentationService action=display-stability-point phase=ReadyShowRequested, workbook=" + _excelInteropService.GetWorkbookFullName (workbook) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_logger.Info ("ShowCreatedCase task pane ready-show requested. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				NewCaseVisibilityObservation.Log (_logger, _excelInteropService, null, workbook, null, "ready-show-requested", "KernelCasePresentationService.ExecuteDeferredPresentationEnhancements", _excelInteropService.GetWorkbookFullName (workbook));
				try {
					MoveInitialCursorToHomeCell (workbook);
					_logger.Info ("ShowCreatedCase cursor positioned. elapsedMs=" + stopwatch.ElapsedMilliseconds);
				} catch (Exception exception2) {
					presentationResult.MarkDegraded ("CursorPositioningSkipped:" + exception2.GetType ().Name);
					_logger.Warn ("ShowCreatedCase cursor positioning skipped after deferred presentation. message=" + exception2.Message + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				}
				_logger.Info ("[KernelFlickerTrace] source=KernelCasePresentationService action=display-stability-point phase=DeferredPresentationCompleted, workbook=" + _excelInteropService.GetWorkbookFullName (workbook) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_logger.Info ("ShowCreatedCase deferred presentation completed. elapsedMs=" + stopwatch.ElapsedMilliseconds);
			} catch (Exception exception) {
				_logger.Error ("ShowCreatedCase deferred presentation failed.", exception);
				presentationResult.MarkDegraded ("DeferredPresentationException:" + exception.GetType ().Name);
			} finally {
				if (!flag) {
					try {
						_transientPaneSuppressionService.ReleaseWorkbook (workbook, "KernelCasePresentationService.ShowCreatedCase.DeferredCleanup");
					} catch (Exception exception2) {
						_logger.Error ("ShowCreatedCase deferred cleanup failed.", exception2);
					}
				}
			}
			return presentationResult;
		}

		private WorkbookWindowVisibilityEnsureResult EnsureWorkbookWindowVisibleBeforeReadyShow (Workbook workbook, Stopwatch stopwatch)
		{
			if (workbook == null) {
				return null;
			}
			WorkbookWindowVisibilityEnsureResult result = _workbookWindowVisibilityService.EnsureVisible (workbook, "KernelCasePresentationService.EnsureWorkbookWindowVisibleBeforeReadyShow");
			switch (result.Outcome) {
			case WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible:
				NewCaseDefaultTimingLogHelper.LogHiddenOpenToWindowVisible (_logger, result.WorkbookFullName, "alreadyVisible");
				_logger.Info ("ShowCreatedCase workbook window visibility ensure skipped because workbook window is already visible. workbook=" + result.WorkbookFullName + ", windowHwnd=" + result.WindowHwnd + ", elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
				break;
			case WorkbookWindowVisibilityEnsureOutcome.MadeVisible:
				NewCaseDefaultTimingLogHelper.LogHiddenOpenToWindowVisible (_logger, result.WorkbookFullName, "madeVisible");
				_logger.Info ("ShowCreatedCase workbook window made visible before ready-show. workbook=" + result.WorkbookFullName + ", windowHwnd=" + result.WindowHwnd + ", elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
				break;
			}
			return result;
		}

		private static void ApplyPresentationOutcome (KernelCaseCreationResult result, CasePresentationResult presentationResult)
		{
			if (result == null || presentationResult == null) {
				return;
			}
			result.PresentationOutcome = presentationResult.Outcome;
			result.PresentationOutcomeReason = presentationResult.OutcomeReason;
		}

		private static void MarkVisibilityDegradationIfNeeded (CasePresentationResult presentationResult, string phase, WorkbookWindowVisibilityEnsureResult ensureResult)
		{
			if (presentationResult == null || IsSuccessfulVisibilityEnsure (ensureResult)) {
				return;
			}
			presentationResult.MarkDegraded ((phase ?? "Visibility") + ":" + ((ensureResult == null) ? "NoResult" : ensureResult.Outcome.ToString ()));
		}

		private static bool IsSuccessfulVisibilityEnsure (WorkbookWindowVisibilityEnsureResult ensureResult)
		{
			return ensureResult != null
				&& (ensureResult.Outcome == WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible || ensureResult.Outcome == WorkbookWindowVisibilityEnsureOutcome.MadeVisible);
		}

		private static string SanitizeOutcomeReason (string outcomeReason)
		{
			return (outcomeReason ?? string.Empty).Replace ("\r\n", " | ").Replace ("\n", " | ").Replace ("\r", " | ");
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
				// 一時的に解決した初期カーソル Range はここで完全解放する既存方針を維持する。
				CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease (range);
			}
		}

		private Range ResolveInitialCursorRange (Workbook caseWorkbook, Worksheet homeWorksheet)
		{
			if (caseWorkbook == null || homeWorksheet == null) {
				return null;
			}
			KernelWorkbookAccessResult kernelAccess = _kernelWorkbookResolverService.ResolveOrOpenReadOnly (caseWorkbook);
			Workbook workbook = kernelAccess.Workbook;
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
				try {
					// Initial cursor lookup only borrows Kernel definitions; close any temporary access here.
					kernelAccess.CloseIfOwned ("KernelCasePresentationService.ResolveInitialCursorRange");
				} catch (Exception exception) {
					_logger.Error ("ResolveInitialCursorRange temporary kernel close failed.", exception);
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

	internal sealed class CasePresentationResult
	{
		private readonly List<string> _outcomeReasons = new List<string> ();

		private CasePresentationResult (CasePresentationOutcome outcome, string outcomeReason)
		{
			Outcome = outcome;
			AddReason (outcomeReason);
		}

		internal CasePresentationOutcome Outcome { get; private set; }

		internal string OutcomeReason
		{
			get
			{
				return _outcomeReasons.Count == 0 ? Outcome.ToString () : string.Join ("|", _outcomeReasons.ToArray ());
			}
		}

		internal static CasePresentationResult Completed ()
		{
			return new CasePresentationResult (CasePresentationOutcome.Completed, null);
		}

		internal static CasePresentationResult Failed (string outcomeReason)
		{
			return new CasePresentationResult (CasePresentationOutcome.Failed, outcomeReason);
		}

		internal void MarkDegraded (string outcomeReason)
		{
			if (Outcome != CasePresentationOutcome.Failed) {
				Outcome = CasePresentationOutcome.Degraded;
			}
			AddReason (outcomeReason);
		}

		internal void MarkFailed (string outcomeReason)
		{
			Outcome = CasePresentationOutcome.Failed;
			AddReason (outcomeReason);
		}

		internal void Merge (CasePresentationResult result)
		{
			if (result == null) {
				return;
			}
			if (result.Outcome == CasePresentationOutcome.Failed) {
				Outcome = CasePresentationOutcome.Failed;
			} else if (result.Outcome == CasePresentationOutcome.Degraded && Outcome == CasePresentationOutcome.Completed) {
				Outcome = CasePresentationOutcome.Degraded;
			}
			foreach (string reason in result._outcomeReasons) {
				AddReason (reason);
			}
		}

		private void AddReason (string outcomeReason)
		{
			if (string.IsNullOrWhiteSpace (outcomeReason)) {
				return;
			}
			string text = outcomeReason.Trim ();
			if (_outcomeReasons.Contains (text)) {
				return;
			}
			_outcomeReasons.Add (text);
		}
	}

	internal static class NewCaseDefaultTimingLogHelper
	{
		internal const string PostReleaseReason = ControlFlowReasons.CreatedCasePostRelease;

		private static readonly object SyncRoot = new object ();

		private static readonly Dictionary<string, Session> Sessions = new Dictionary<string, Session> (StringComparer.OrdinalIgnoreCase);

		private sealed class Session
		{
			internal readonly HashSet<string> Keys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);

			internal Stopwatch HiddenOpenToWindowVisibleStopwatch;

			internal Stopwatch TaskPaneReadyWaitToRefreshCompletedStopwatch;

			internal Stopwatch WaitUiCloseToFinalForegroundStableStopwatch;

			internal bool HiddenOpenToWindowVisibleLogged;

			internal bool TaskPaneReadyWaitToRefreshCompletedLogged;

			internal bool WaitUiCloseToFinalForegroundStableLogged;
		}

		internal static void BeginCaseCreation (string workbookPath)
		{
			string text = NormalizeKey (workbookPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value)) {
					value = new Session ();
				}
				RegisterKey (value, text);
			}
		}

		internal static void BeginPresentation (string workbookPath)
		{
			string text = NormalizeKey (workbookPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value)) {
					value = new Session ();
				}
				RegisterKey (value, text);
				if (value.HiddenOpenToWindowVisibleStopwatch == null) {
					value.HiddenOpenToWindowVisibleStopwatch = Stopwatch.StartNew ();
				}
			}
		}

		internal static void AttachWorkbookAlias (string existingWorkbookPath, string aliasWorkbookPath)
		{
			string text = NormalizeKey (existingWorkbookPath);
			string text2 = NormalizeKey (aliasWorkbookPath);
			if (string.IsNullOrWhiteSpace (text) || string.IsNullOrWhiteSpace (text2) || string.Equals (text, text2, StringComparison.OrdinalIgnoreCase)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value)) {
					return;
				}
				RegisterKey (value, text2);
			}
		}

		internal static void StartTaskPaneReadyWait (string workbookPath)
		{
			string text = NormalizeKey (workbookPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value) || value.TaskPaneReadyWaitToRefreshCompletedStopwatch != null) {
					return;
				}
				value.TaskPaneReadyWaitToRefreshCompletedStopwatch = Stopwatch.StartNew ();
			}
		}

		internal static void StartWaitUiCloseToFinalForegroundStable (string workbookPath)
		{
			string text = NormalizeKey (workbookPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value) || value.WaitUiCloseToFinalForegroundStableStopwatch != null) {
					return;
				}
				value.WaitUiCloseToFinalForegroundStableStopwatch = Stopwatch.StartNew ();
			}
		}

		internal static void LogHiddenOpenToWindowVisible (Logger logger, string workbookPath, string outcome)
		{
			string text = NormalizeKey (workbookPath);
			if (logger == null || string.IsNullOrWhiteSpace (text)) {
				return;
			}
			long elapsedMilliseconds;
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value) || value.HiddenOpenToWindowVisibleLogged || value.HiddenOpenToWindowVisibleStopwatch == null) {
					return;
				}
				value.HiddenOpenToWindowVisibleLogged = true;
				elapsedMilliseconds = value.HiddenOpenToWindowVisibleStopwatch.ElapsedMilliseconds;
			}
			logger.Info ("NewCaseDefault timing. segment=hiddenOpenToWindowVisible, caseWorkbookPath=" + text + ", outcome=" + (outcome ?? string.Empty) + ", elapsedMs=" + elapsedMilliseconds);
		}

		internal static void LogDetail (Logger logger, string workbookPath, string segment, string phase, long elapsedMilliseconds, string details = "")
		{
			string text = NormalizeKey (workbookPath);
			if (logger == null || string.IsNullOrWhiteSpace (text)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.ContainsKey (text)) {
					return;
				}
			}
			string text2 = string.IsNullOrWhiteSpace (details) ? string.Empty : ", " + details.Trim ();
			logger.Info ("NewCaseDefault timing detail. segment=" + (segment ?? string.Empty) + ", phase=" + (phase ?? string.Empty) + ", caseWorkbookPath=" + text + text2 + ", elapsedMs=" + Math.Max (0L, elapsedMilliseconds));
		}

		internal static void LogTaskPaneReadyWaitToRefreshCompleted (Logger logger, string workbookPath, string reason, bool refreshed, string completion)
		{
			string text = NormalizeKey (workbookPath);
			if (logger == null || string.IsNullOrWhiteSpace (text) || !string.Equals (reason, PostReleaseReason, StringComparison.Ordinal)) {
				return;
			}
			long elapsedMilliseconds;
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value) || value.TaskPaneReadyWaitToRefreshCompletedLogged || value.TaskPaneReadyWaitToRefreshCompletedStopwatch == null) {
					return;
				}
				value.TaskPaneReadyWaitToRefreshCompletedLogged = true;
				elapsedMilliseconds = value.TaskPaneReadyWaitToRefreshCompletedStopwatch.ElapsedMilliseconds;
			}
			logger.Info ("NewCaseDefault timing. segment=taskPaneReadyWaitToRefreshCompleted, caseWorkbookPath=" + text + ", completion=" + (completion ?? string.Empty) + ", refreshed=" + refreshed + ", elapsedMs=" + elapsedMilliseconds);
		}

		internal static void LogWaitUiCloseToFinalForegroundStable (Logger logger, string workbookPath, string reason, bool recovered)
		{
			if (!string.Equals (reason, PostReleaseReason, StringComparison.Ordinal)) {
				return;
			}
			LogWaitUiCloseToFinalForegroundStableCore (logger, workbookPath, recovered, "postReleaseForegroundRecovery");
		}

		internal static void LogWaitUiCloseToFinalForegroundStableIfPending (Logger logger, string workbookPath, string outcome)
		{
			LogWaitUiCloseToFinalForegroundStableCore (logger, workbookPath, recovered: false, outcome ?? string.Empty);
		}

		private static void LogWaitUiCloseToFinalForegroundStableCore (Logger logger, string workbookPath, bool recovered, string outcome)
		{
			string text = NormalizeKey (workbookPath);
			if (logger == null || string.IsNullOrWhiteSpace (text)) {
				return;
			}
			long elapsedMilliseconds;
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value) || value.WaitUiCloseToFinalForegroundStableLogged || value.WaitUiCloseToFinalForegroundStableStopwatch == null) {
					return;
				}
				value.WaitUiCloseToFinalForegroundStableLogged = true;
				elapsedMilliseconds = value.WaitUiCloseToFinalForegroundStableStopwatch.ElapsedMilliseconds;
				RemoveSession (value);
			}
			logger.Info ("NewCaseDefault timing. segment=waitUiCloseToFinalForegroundStable, caseWorkbookPath=" + text + ", outcome=" + (outcome ?? string.Empty) + ", recovered=" + recovered + ", elapsedMs=" + elapsedMilliseconds);
		}

		internal static void Clear (string workbookPath)
		{
			string text = NormalizeKey (workbookPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return;
			}
			lock (SyncRoot) {
				if (!Sessions.TryGetValue (text, out Session value)) {
					return;
				}
				RemoveSession (value);
			}
		}

		private static string NormalizeKey (string workbookPath)
		{
			return (workbookPath ?? string.Empty).Trim ();
		}

		private static void RegisterKey (Session session, string key)
		{
			if (session == null || string.IsNullOrWhiteSpace (key)) {
				return;
			}
			session.Keys.Add (key);
			Sessions [key] = session;
		}

		private static void RemoveSession (Session session)
		{
			if (session == null) {
				return;
			}
			foreach (string key in session.Keys) {
				Sessions.Remove (key);
			}
			session.Keys.Clear ();
		}
	}
}
