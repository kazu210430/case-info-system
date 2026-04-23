using System;
using System.Diagnostics;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelCaseCreationService
	{
		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly KernelCasePathService _kernelCasePathService;

		private readonly CaseWorkbookInitializer _caseWorkbookInitializer;

		private readonly CaseWorkbookOpenStrategy _caseWorkbookOpenStrategy;

		private readonly TransientPaneSuppressionService _transientPaneSuppressionService;

		private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

		private readonly ExcelInteropService _excelInteropService;

		private readonly Logger _logger;

		internal KernelCaseCreationService (KernelWorkbookService kernelWorkbookService, KernelCasePathService kernelCasePathService, CaseWorkbookInitializer caseWorkbookInitializer, CaseWorkbookOpenStrategy caseWorkbookOpenStrategy, TransientPaneSuppressionService transientPaneSuppressionService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, ExcelInteropService excelInteropService, Logger logger)
		{
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_kernelCasePathService = kernelCasePathService ?? throw new ArgumentNullException ("kernelCasePathService");
			_caseWorkbookInitializer = caseWorkbookInitializer ?? throw new ArgumentNullException ("caseWorkbookInitializer");
			_caseWorkbookOpenStrategy = caseWorkbookOpenStrategy ?? throw new ArgumentNullException ("caseWorkbookOpenStrategy");
			_transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException ("transientPaneSuppressionService");
			_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal KernelCaseCreationResult CreateCase (KernelCaseCreationRequest request)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			Workbook openKernelWorkbook = _kernelWorkbookService.GetOpenKernelWorkbook ();
			if (openKernelWorkbook == null) {
				throw new InvalidOperationException ("Kernel workbook is not open.");
			}
			KernelCaseCreationPlan kernelCaseCreationPlan = ResolveCreationPlan (openKernelWorkbook, request, stopwatch);
			_logger.Info ("Kernel case creation plan resolved. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds + ", casePath=" + kernelCaseCreationPlan.CaseWorkbookPath);
			File.Copy (kernelCaseCreationPlan.BaseWorkbookPath, kernelCaseCreationPlan.CaseWorkbookPath, overwrite: false);
			_logger.Info ("Kernel case base workbook copied. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			return CreateSavedCase (openKernelWorkbook, kernelCaseCreationPlan);
		}

		private KernelCaseCreationPlan ResolveCreationPlan (Workbook kernelWorkbook, KernelCaseCreationRequest request, Stopwatch outerStopwatch)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			string text = _kernelCasePathService.ResolveSystemRoot (kernelWorkbook);
			if (string.IsNullOrWhiteSpace (text)) {
				throw new InvalidOperationException ("Kernel SYSTEM_ROOT could not be resolved.");
			}
			string text2 = _kernelCasePathService.ResolveBaseWorkbookPath (text);
			if (!File.Exists (text2)) {
				throw new FileNotFoundException ("Base workbook was not found.", text2);
			}
			string text3 = KernelNamingService.NormalizeNameRuleA (_excelInteropService.TryGetDocumentProperty (kernelWorkbook, "NAME_RULE_A"));
			string nameRuleB = KernelNamingService.NormalizeNameRuleB (_excelInteropService.TryGetDocumentProperty (kernelWorkbook, "NAME_RULE_B"));
			string customerName = (request.CustomerName ?? string.Empty).Trim ();
			string folderName = KernelNamingService.BuildFolderName (text3, customerName, DateTime.Today);
			string text4 = _kernelCasePathService.ResolveCaseFolderPath (request, folderName);
			if (string.IsNullOrWhiteSpace (text4)) {
				throw new InvalidOperationException ("CASE folder path could not be resolved.");
			}
			if (!_kernelCasePathService.EnsureFolderExists (text4)) {
				throw new InvalidOperationException ("CASE folder could not be created.");
			}
			string extension = _kernelCasePathService.ResolveCaseWorkbookExtension (text2);
			string caseWorkbookName = BuildCaseWorkbookName (customerName, extension);
			string caseWorkbookPath = _kernelCasePathService.BuildCaseWorkbookPath (text4, caseWorkbookName);
			_logger.Info ("Kernel case creation plan built. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds + ", folderPath=" + text4);
			return new KernelCaseCreationPlan {
				Mode = request.Mode,
				CustomerName = customerName,
				SystemRoot = text,
				BaseWorkbookPath = text2,
				CaseFolderPath = text4,
				CaseWorkbookPath = caseWorkbookPath,
				NameRuleA = text3,
				NameRuleB = nameRuleB
			};
		}

		private static string BuildCaseWorkbookName (string customerName, string extension)
		{
			return KernelNamingService.BuildCaseBookName (customerName, extension);
		}

		internal KernelCaseCreationResult CreateSavedCase (Workbook kernelWorkbook, KernelCaseCreationPlan plan)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			Workbook workbook = null;
			Application application = null;
			bool previousDisplayAlerts = true;
			bool transientSuppressionRegistered = false;
			try {
				application = (kernelWorkbook == null) ? null : kernelWorkbook.Application;
				if (application == null) {
					throw new InvalidOperationException ("Excel application could not be resolved.");
				}
				previousDisplayAlerts = application.DisplayAlerts;
				_transientPaneSuppressionService.SuppressPath (plan.CaseWorkbookPath, "KernelCaseCreationService.CreateSavedCase");
				transientSuppressionRegistered = true;
				if (ShouldUseHiddenCreateSession (plan.Mode)) {
					CreateSavedCaseWithoutShowing (kernelWorkbook, plan, application, stopwatch);
				} else {
					workbook = application.Workbooks.Open (plan.CaseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
					_logger.Info ("Kernel saved CASE workbook opened. path=" + plan.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
					_caseWorkbookInitializer.InitializeForVisibleCreate (kernelWorkbook, workbook, plan);
					_logger.Info ("Kernel saved CASE initialized. path=" + plan.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
					workbook.Save ();
					_logger.Info ("Kernel saved CASE saved. path=" + plan.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
					using (_caseWorkbookLifecycleService.BeginManagedCloseScope (workbook)) {
						WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose (workbook);
						application.DisplayAlerts = false;
						workbook.Close (false, Type.Missing, Type.Missing);
					}
					workbook = null;
				}
				_transientPaneSuppressionService.ReleasePath (plan.CaseWorkbookPath, "KernelCaseCreationService.CreateSavedCase.Completed");
				transientSuppressionRegistered = false;
				_logger.Info ("Kernel saved CASE workbook closed. path=" + plan.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_logger.Info ("Saved CASE created. path=" + plan.CaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return new KernelCaseCreationResult {
					Success = true,
					Mode = plan.Mode,
					CaseFolderPath = plan.CaseFolderPath,
					CaseWorkbookPath = plan.CaseWorkbookPath,
					UserMessage = string.Empty
				};
			} catch {
				try {
					if (workbook != null) {
						using (_caseWorkbookLifecycleService.BeginManagedCloseScope (workbook)) {
							WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose (workbook);
							if (application != null) {
								application.DisplayAlerts = false;
							}
							workbook.Close (false, Type.Missing, Type.Missing);
						}
					}
				} catch {
				}
				throw;
			} finally {
				if (transientSuppressionRegistered) {
					try {
						_transientPaneSuppressionService.ReleasePath (plan.CaseWorkbookPath, "KernelCaseCreationService.CreateSavedCase.Finally");
					} catch {
					}
				}
				try {
					if (application != null) {
						application.DisplayAlerts = previousDisplayAlerts;
					}
				} catch {
				}
			}
		}

		private void CreateSavedCaseWithoutShowing (Workbook kernelWorkbook, KernelCaseCreationPlan plan, Application fallbackApplication, Stopwatch stopwatch)
		{
			CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession hiddenCaseWorkbookSession = null;
			Workbook workbook = null;
			Application application = fallbackApplication;
			try {
				hiddenCaseWorkbookSession = _caseWorkbookOpenStrategy.OpenHiddenWorkbook (plan.CaseWorkbookPath);
				if (hiddenCaseWorkbookSession == null || hiddenCaseWorkbookSession.Workbook == null) {
					throw new InvalidOperationException ("CASE workbook hidden session could not be opened.");
				}
				workbook = hiddenCaseWorkbookSession.Workbook;
				application = hiddenCaseWorkbookSession.Application ?? fallbackApplication;
				_logger.Info ("Kernel saved CASE workbook hidden session opened. path=" + plan.CaseWorkbookPath + ", route=" + hiddenCaseWorkbookSession.RouteName + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				_caseWorkbookInitializer.InitializeForHiddenCreate (kernelWorkbook, workbook, plan);
				_logger.Info ("Kernel saved CASE hidden initialized. path=" + plan.CaseWorkbookPath + ", route=" + hiddenCaseWorkbookSession.RouteName + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				workbook.Save ();
				_logger.Info ("Kernel saved CASE hidden saved. path=" + plan.CaseWorkbookPath + ", route=" + hiddenCaseWorkbookSession.RouteName + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				using (_caseWorkbookLifecycleService.BeginManagedCloseScope (workbook)) {
					WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose (workbook);
					if (application != null) {
						application.DisplayAlerts = false;
					}
					hiddenCaseWorkbookSession.Close ();
				}
				_logger.Info ("Kernel saved CASE hidden session closed. path=" + plan.CaseWorkbookPath + ", route=" + hiddenCaseWorkbookSession.RouteName + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			} catch {
				if (hiddenCaseWorkbookSession != null && workbook != null) {
					try {
						using (_caseWorkbookLifecycleService.BeginManagedCloseScope (workbook)) {
							WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose (workbook);
							if (application != null) {
								application.DisplayAlerts = false;
							}
							hiddenCaseWorkbookSession.Abort ();
						}
					} catch {
					}
				}
				throw;
			}
		}

		private static bool ShouldUseHiddenCreateSession (KernelCaseCreationMode mode)
		{
			return mode == KernelCaseCreationMode.CreateCaseBatch;
		}

		private static string SafeApplicationHwnd (Application application)
		{
			try {
				return (application == null) ? string.Empty : (Convert.ToString (application.Hwnd) ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private void LogCreateSavedCasePhase (string phaseName, string path, string routeName, Stopwatch stopwatch, long phaseStartElapsedMs)
		{
			_logger.Info ("Kernel saved CASE phase completed. phase=" + (phaseName ?? string.Empty) + ", path=" + (path ?? string.Empty) + ", route=" + (routeName ?? string.Empty) + ", phaseElapsedMs=" + Math.Max (0L, stopwatch.ElapsedMilliseconds - phaseStartElapsedMs) + ", totalElapsedMs=" + stopwatch.ElapsedMilliseconds);
		}

		private string PrepareWorkingCaseWorkbookPath (string finalCaseWorkbookPath, string reason, Stopwatch stopwatch)
		{
			// OneDrive など同期配下の CASE だけは、hidden Excel による初期化中のフリーズ回避のため
			// 一時的にローカル作業コピーへ退避する。適用範囲は CASE 初期化中のみで、表示・運用は final path に戻す。
			if (!_kernelCasePathService.IsUnderSyncRoot (finalCaseWorkbookPath)) {
				return finalCaseWorkbookPath;
			}
			string text = _kernelCasePathService.BuildLocalWorkingCaseWorkbookPath (finalCaseWorkbookPath);
			if (string.IsNullOrWhiteSpace (text)) {
				_logger.Info ("Kernel case local working path was not prepared. reason=" + reason + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return finalCaseWorkbookPath;
			}
			File.Copy (finalCaseWorkbookPath, text, overwrite: false);
			_logger.Info ("Kernel case local working copy prepared. reason=" + reason + ", localPath=" + text + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			return text;
		}

		private void FinalizeWorkingCaseWorkbookPath (string workingCaseWorkbookPath, string finalCaseWorkbookPath, string reason, Stopwatch stopwatch)
		{
			// 作業コピーは初期化が終わった時点で final path へ戻し、CASE の実体を temp に残さない。
			if (!string.Equals (workingCaseWorkbookPath, finalCaseWorkbookPath, StringComparison.OrdinalIgnoreCase)) {
				if (!_kernelCasePathService.MoveLocalWorkingCaseToFinalPath (workingCaseWorkbookPath, finalCaseWorkbookPath)) {
					throw new IOException ("Initialized CASE workbook could not be moved to final path.");
				}
				_logger.Info ("Kernel case local working copy finalized. reason=" + reason + ", finalPath=" + finalCaseWorkbookPath + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			}
		}

	}
}
