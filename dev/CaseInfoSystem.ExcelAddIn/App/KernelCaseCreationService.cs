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

		private readonly FolderWindowService _folderWindowService;

		private readonly TransientPaneSuppressionService _transientPaneSuppressionService;

		private readonly ExcelInteropService _excelInteropService;

		private readonly Logger _logger;

		internal KernelCaseCreationService (KernelWorkbookService kernelWorkbookService, KernelCasePathService kernelCasePathService, CaseWorkbookInitializer caseWorkbookInitializer, CaseWorkbookOpenStrategy caseWorkbookOpenStrategy, FolderWindowService folderWindowService, TransientPaneSuppressionService transientPaneSuppressionService, ExcelInteropService excelInteropService, Logger logger)
		{
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_kernelCasePathService = kernelCasePathService ?? throw new ArgumentNullException ("kernelCasePathService");
			_caseWorkbookInitializer = caseWorkbookInitializer ?? throw new ArgumentNullException ("caseWorkbookInitializer");
			_caseWorkbookOpenStrategy = caseWorkbookOpenStrategy ?? throw new ArgumentNullException ("caseWorkbookOpenStrategy");
			_folderWindowService = folderWindowService ?? throw new ArgumentNullException ("folderWindowService");
			_transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException ("transientPaneSuppressionService");
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
			OpenCaseFolderEarly (text4, request.Mode, outerStopwatch);
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
			CaseWorkbookOpenStrategy.HiddenCaseWorkbookSession hiddenCaseWorkbookSession = null;
			string text = PrepareWorkingCaseWorkbookPath (plan.CaseWorkbookPath, "CreateSavedCase", stopwatch);
			try {
				hiddenCaseWorkbookSession = _caseWorkbookOpenStrategy.OpenHiddenWorkbook (text);
				_logger.Info ("Kernel saved CASE session ready. path=" + text + ", appHwnd=" + SafeApplicationHwnd (hiddenCaseWorkbookSession.Application));
				_caseWorkbookInitializer.InitializeForHiddenCreate (kernelWorkbook, hiddenCaseWorkbookSession.Workbook, plan);
				_logger.Info ("Kernel saved CASE initialized. path=" + text + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				hiddenCaseWorkbookSession.Workbook.Save ();
				_logger.Info ("Kernel saved CASE saved. path=" + text + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				hiddenCaseWorkbookSession.Workbook.Close (false, Type.Missing, Type.Missing);
				_logger.Info ("Kernel saved CASE session quitting. path=" + text + ", appHwnd=" + SafeApplicationHwnd (hiddenCaseWorkbookSession.Application));
				hiddenCaseWorkbookSession.Application.Quit ();
				hiddenCaseWorkbookSession = null;
				FinalizeWorkingCaseWorkbookPath (text, plan.CaseWorkbookPath, "CreateSavedCase", stopwatch);
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
					if (hiddenCaseWorkbookSession != null) {
						hiddenCaseWorkbookSession.Workbook.Close (false, Type.Missing, Type.Missing);
						_logger.Info ("Kernel saved CASE session quitting after failure. path=" + text + ", appHwnd=" + SafeApplicationHwnd (hiddenCaseWorkbookSession.Application));
						hiddenCaseWorkbookSession.Application.Quit ();
					}
				} catch {
				}
				throw;
			}
		}

		private static string SafeApplicationHwnd (Application application)
		{
			try {
				return (application == null) ? string.Empty : (Convert.ToString (application.Hwnd) ?? string.Empty);
			} catch {
				return string.Empty;
			}
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

		private void OpenCaseFolderEarly (string caseFolderPath, KernelCaseCreationMode mode, Stopwatch stopwatch)
		{
			if (string.IsNullOrWhiteSpace (caseFolderPath)) {
				return;
			}
			try {
				_folderWindowService.OpenFolder (caseFolderPath, "KernelCaseCreationService.OpenCaseFolderEarly");
				_logger.Info ("Kernel case folder opened before save completed. mode=" + mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			} catch (Exception exception) {
				_logger.Error ("OpenCaseFolderEarly failed.", exception);
			}
		}
	}
}
