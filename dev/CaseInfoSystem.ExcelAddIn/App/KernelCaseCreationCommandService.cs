using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelCaseCreationCommandService
	{
		private const string CreateCaseFailedMessage = "案件作成に失敗しました。";

		private const string OpenCaseFailedMessage = "保存は完了しましたが、案件情報を開けませんでした。";

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly KernelCaseCreationService _kernelCaseCreationService;

		private readonly KernelCasePathService _kernelCasePathService;

		private readonly KernelCasePresentationService _kernelCasePresentationService;

		private readonly CreatedCaseOpenPromptService _createdCaseOpenPromptService;

		private readonly CreatedCasePresentationWaitService _createdCasePresentationWaitService;

		private readonly ExcelInteropService _excelInteropService;

		private readonly Logger _logger;

		internal KernelCaseCreationCommandService (KernelWorkbookService kernelWorkbookService, KernelCaseCreationService kernelCaseCreationService, KernelCasePathService kernelCasePathService, KernelCasePresentationService kernelCasePresentationService, CreatedCaseOpenPromptService createdCaseOpenPromptService, CreatedCasePresentationWaitService createdCasePresentationWaitService, ExcelInteropService excelInteropService, Logger logger)
		{
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_kernelCaseCreationService = kernelCaseCreationService ?? throw new ArgumentNullException ("kernelCaseCreationService");
			_kernelCasePathService = kernelCasePathService ?? throw new ArgumentNullException ("kernelCasePathService");
			_kernelCasePresentationService = kernelCasePresentationService ?? throw new ArgumentNullException ("kernelCasePresentationService");
			_createdCaseOpenPromptService = createdCaseOpenPromptService ?? throw new ArgumentNullException ("createdCaseOpenPromptService");
			_createdCasePresentationWaitService = createdCasePresentationWaitService ?? throw new ArgumentNullException ("createdCasePresentationWaitService");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal KernelCaseCreationResult ExecuteNewCaseDefault (string customerName)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			Workbook workbook = RequireKernelWorkbook ();
			string text = _excelInteropService.TryGetDocumentProperty (workbook, "DEFAULT_ROOT");
			_logger.Info ("Kernel case command start. mode=NewCaseDefault, hasDefaultRoot=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				string text2 = _kernelCasePathService.SelectFolderPath ("既定のフォルダを選択してください。", string.Empty);
				if (string.IsNullOrWhiteSpace (text2)) {
					return BuildFailure (string.Empty);
				}
				_excelInteropService.SetDocumentProperty (workbook, "DEFAULT_ROOT", text2);
				workbook.Save ();
				text = text2;
				_logger.Info ("Kernel case command default root saved. mode=NewCaseDefault, elapsedMs=" + stopwatch.ElapsedMilliseconds);
			}
			return Execute (new KernelCaseCreationRequest {
				CustomerName = customerName,
				Mode = KernelCaseCreationMode.NewCaseDefault,
				DefaultRoot = text
			});
		}

		internal KernelCaseCreationResult ExecuteCreateCaseSingle (string customerName)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			string text = SelectFolderAndRemember ();
			_logger.Info ("Kernel case command folder selected. mode=CreateCaseSingle, selected=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				return BuildFailure (string.Empty);
			}
			return Execute (new KernelCaseCreationRequest {
				CustomerName = customerName,
				Mode = KernelCaseCreationMode.CreateCaseSingle,
				SelectedFolderPath = text
			});
		}

		internal KernelCaseCreationResult ExecuteCreateCaseBatch (string customerName)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			string text = SelectFolderAndRemember ();
			_logger.Info ("Kernel case command folder selected. mode=CreateCaseBatch, selected=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				return BuildFailure (string.Empty);
			}
			return Execute (new KernelCaseCreationRequest {
				CustomerName = customerName,
				Mode = KernelCaseCreationMode.CreateCaseBatch,
				SelectedFolderPath = text
			});
		}

		private KernelCaseCreationResult Execute (KernelCaseCreationRequest request)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			CreatedCasePresentationWaitService.WaitSession waitSession = null;
			long waitUiShownElapsedMs = -1L;
			bool waitSessionTransferred = false;
			try {
				ValidateRequest (request);
				_logger.Info ("Kernel case command validated. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				if (ShouldPromptToOpenCreatedCase (request.Mode)) {
					waitSession = _createdCasePresentationWaitService.ShowWaiting (stopwatch);
					waitSession.UpdateStage (CreatedCasePresentationWaitService.CreatingStageTitle);
					if (request.Mode == KernelCaseCreationMode.NewCaseDefault) {
						waitUiShownElapsedMs = stopwatch.ElapsedMilliseconds;
					}
				}
				KernelCaseCreationResult kernelCaseCreationResult = _kernelCaseCreationService.CreateCase (request);
				if (!kernelCaseCreationResult.Success) {
					return kernelCaseCreationResult;
				}
				if (request.Mode == KernelCaseCreationMode.NewCaseDefault && waitUiShownElapsedMs >= 0L) {
					_logger.Info ("NewCaseDefault timing. segment=waitUiShownToCaseCreated, caseWorkbookPath=" + (kernelCaseCreationResult.CaseWorkbookPath ?? string.Empty) + ", elapsedMs=" + Math.Max (0L, stopwatch.ElapsedMilliseconds - waitUiShownElapsedMs));
				}
				if (!ShouldPromptToOpenCreatedCase (kernelCaseCreationResult.Mode)) {
					PresentCaseFolderBestEffort (kernelCaseCreationResult, "KernelCaseCreationCommandService.Execute.NoPrompt");
					kernelCaseCreationResult.ShouldCloseKernelHome = false;
					_logger.Info ("Kernel case command completed without open prompt. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
					return kernelCaseCreationResult;
				}
				TryStartCaseFolderEarlyOpen (kernelCaseCreationResult, stopwatch);
				waitSessionTransferred = true;
				KernelCaseCreationResult kernelCaseCreationResult2 = CompleteInteractiveOpenFlow (kernelCaseCreationResult, stopwatch, waitSession);
				waitSession = null;
				return kernelCaseCreationResult2;
			} catch (Exception exception) {
				_logger.Error ("Kernel case creation failed.", exception);
				return BuildFailure ("案件作成に失敗しました。");
			}
			finally {
				if (!waitSessionTransferred) {
					CloseWaitSession (waitSession, restoreOwner: true);
				}
			}
		}

		private KernelCaseCreationResult CompleteInteractiveOpenFlow (KernelCaseCreationResult result, Stopwatch stopwatch, CreatedCasePresentationWaitService.WaitSession waitSession)
		{
			if (result == null) {
				throw new ArgumentNullException ("result");
			}
			_logger.Info ("Kernel case auto-open selected. mode=" + result.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			try {
				KernelCaseCreationResult kernelCaseCreationResult = _kernelCasePresentationService.OpenCreatedCase (result, waitSession);
				kernelCaseCreationResult.ShouldCloseKernelHome = true;
				_logger.Info ("Kernel case presentation complete. mode=" + kernelCaseCreationResult.Mode.ToString () + ", success=" + kernelCaseCreationResult.Success + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return kernelCaseCreationResult;
			} catch (Exception exception) {
				_logger.Error ("Kernel case open after save failed.", exception);
				return BuildFailure ("保存は完了しましたが、案件情報を開けませんでした。");
			}
		}

		private void TryStartCaseFolderEarlyOpen (KernelCaseCreationResult result, Stopwatch stopwatch)
		{
			if (!ShouldPresentCaseFolder (result)) {
				return;
			}
			if (!ShouldStartCaseFolderEarlyOpen (result.Mode)) {
				_logger.Info ("Kernel case early folder open suppressed. mode=" + result.Mode.ToString () + ", folderPath=" + result.CaseFolderPath + ", elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
				return;
			}
			_kernelCasePresentationService.OpenCaseFolder (result.CaseFolderPath, "KernelCaseCreationCommandService.Execute.EarlyPreOpen");
			_logger.Info ("Kernel case early folder open requested. mode=" + result.Mode.ToString () + ", folderPath=" + result.CaseFolderPath + ", elapsedMs=" + ((stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds));
		}

		private static bool ShouldPromptToOpenCreatedCase (KernelCaseCreationMode mode)
		{
			return mode == KernelCaseCreationMode.NewCaseDefault || mode == KernelCaseCreationMode.CreateCaseSingle;
		}

		private static bool ShouldStartCaseFolderEarlyOpen (KernelCaseCreationMode mode)
		{
			return false;
		}

		private void PresentCaseFolderBestEffort (KernelCaseCreationResult result, string reason)
		{
			if (!ShouldPresentCaseFolder (result)) {
				return;
			}
			if (ShouldWaitForCaseFolderPresentation (result.Mode)) {
				_kernelCasePresentationService.OpenCaseFolderAndWait (result.CaseFolderPath, reason);
				return;
			}
			_kernelCasePresentationService.OpenCaseFolder (result.CaseFolderPath, reason);
		}

		private static bool ShouldPresentCaseFolder (KernelCaseCreationResult result)
		{
			return result != null && result.Success && !string.IsNullOrWhiteSpace (result.CaseFolderPath);
		}

		private static bool ShouldWaitForCaseFolderPresentation (KernelCaseCreationMode mode)
		{
			return mode == KernelCaseCreationMode.CreateCaseBatch;
		}

		private static void ValidateRequest (KernelCaseCreationRequest request)
		{
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			if (request.Mode == KernelCaseCreationMode.NewCaseDefault && string.IsNullOrWhiteSpace (request.DefaultRoot)) {
				throw new InvalidOperationException ("DEFAULT_ROOT が未設定です。");
			}
			if ((request.Mode == KernelCaseCreationMode.CreateCaseSingle || request.Mode == KernelCaseCreationMode.CreateCaseBatch) && string.IsNullOrWhiteSpace (request.SelectedFolderPath)) {
				throw new InvalidOperationException ("保存先フォルダが未選択です。");
			}
		}

		private Workbook RequireKernelWorkbook ()
		{
			Workbook openKernelWorkbook = _kernelWorkbookService.GetOpenKernelWorkbook ();
			if (openKernelWorkbook == null) {
				throw new InvalidOperationException ("Kernel ブックを取得できませんでした。");
			}
			return openKernelWorkbook;
		}

		private string SelectFolderAndRemember ()
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			Workbook workbook = RequireKernelWorkbook ();
			string initialDirectory = _excelInteropService.TryGetDocumentProperty (workbook, "LAST_PICK_FOLDER");
			string text = _kernelCasePathService.SelectFolderPath ("保存先フォルダを選択してください。", initialDirectory);
			_logger.Info ("Kernel case folder dialog completed. selected=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				return string.Empty;
			}
			_excelInteropService.SetDocumentProperty (workbook, "LAST_PICK_FOLDER", text);
			workbook.Save ();
			_logger.Info ("Kernel case folder remembered. elapsedMs=" + stopwatch.ElapsedMilliseconds);
			return text;
		}

		private static KernelCaseCreationResult BuildFailure (string userMessage)
		{
			return new KernelCaseCreationResult {
				Success = false,
				UserMessage = (userMessage ?? string.Empty),
				ShouldCloseKernelHome = false
			};
		}

		private static void CloseWaitSession (CreatedCasePresentationWaitService.WaitSession waitSession, bool restoreOwner)
		{
			if (waitSession == null) {
				return;
			}
			if (restoreOwner) {
				waitSession.CloseAndRestoreOwner ();
			} else {
				waitSession.CloseForSuccessfulPresentation ();
			}
			waitSession.Dispose ();
		}
	}
}
