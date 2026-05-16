using System;
using System.Diagnostics;
using System.IO;
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

		private readonly CreatedCasePresentationWaitService _createdCasePresentationWaitService;

		private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;

		private readonly ExcelInteropService _excelInteropService;

		private readonly Logger _logger;

		internal KernelCaseCreationCommandService (KernelWorkbookService kernelWorkbookService, KernelCaseCreationService kernelCaseCreationService, KernelCasePathService kernelCasePathService, KernelCasePresentationService kernelCasePresentationService, CreatedCasePresentationWaitService createdCasePresentationWaitService, CaseWorkbookLifecycleService caseWorkbookLifecycleService, ExcelInteropService excelInteropService, Logger logger)
		{
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_kernelCaseCreationService = kernelCaseCreationService ?? throw new ArgumentNullException ("kernelCaseCreationService");
			_kernelCasePathService = kernelCasePathService ?? throw new ArgumentNullException ("kernelCasePathService");
			_kernelCasePresentationService = kernelCasePresentationService ?? throw new ArgumentNullException ("kernelCasePresentationService");
			_createdCasePresentationWaitService = createdCasePresentationWaitService ?? throw new ArgumentNullException ("createdCasePresentationWaitService");
			_caseWorkbookLifecycleService = caseWorkbookLifecycleService ?? throw new ArgumentNullException ("caseWorkbookLifecycleService");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal KernelCaseCreationResult ExecuteNewCaseDefault (Workbook kernelWorkbook, string expectedSystemRoot, string customerName)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			if (!TryValidateBoundKernelWorkbook (kernelWorkbook, expectedSystemRoot, "ExecuteNewCaseDefault")) {
				return BuildFailure (CreateCaseFailedMessage);
			}
			string text = _excelInteropService.TryGetDocumentProperty (kernelWorkbook, "DEFAULT_ROOT");
			_logger.Info ("Kernel case command start. mode=NewCaseDefault, hasDefaultRoot=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				string text2 = _kernelCasePathService.SelectFolderPath ("既定のフォルダを選択してください。", string.Empty);
				if (string.IsNullOrWhiteSpace (text2)) {
					return BuildFailure (string.Empty);
				}
				_excelInteropService.SetDocumentProperty (kernelWorkbook, "DEFAULT_ROOT", text2);
				kernelWorkbook.Save ();
				text = text2;
				_logger.Info ("Kernel case command default root saved. mode=NewCaseDefault, elapsedMs=" + stopwatch.ElapsedMilliseconds);
			}
			return Execute (new KernelCaseCreationRequest {
				CustomerName = customerName,
				Mode = KernelCaseCreationMode.NewCaseDefault,
				DefaultRoot = text
			}, kernelWorkbook, expectedSystemRoot);
		}

		internal KernelCaseCreationResult ExecuteCreateCaseSingle (Workbook kernelWorkbook, string expectedSystemRoot, string customerName)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			string text = SelectFolderAndRemember (kernelWorkbook, expectedSystemRoot);
			_logger.Info ("Kernel case command folder selected. mode=CreateCaseSingle, selected=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				return BuildFailure (string.Empty);
			}
			return Execute (new KernelCaseCreationRequest {
				CustomerName = customerName,
				Mode = KernelCaseCreationMode.CreateCaseSingle,
				SelectedFolderPath = text
			}, kernelWorkbook, expectedSystemRoot);
		}

		internal KernelCaseCreationResult ExecuteCreateCaseBatch (Workbook kernelWorkbook, string expectedSystemRoot, string customerName)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			string text = SelectFolderAndRemember (kernelWorkbook, expectedSystemRoot);
			_logger.Info ("Kernel case command folder selected. mode=CreateCaseBatch, selected=" + !string.IsNullOrWhiteSpace (text) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
			if (string.IsNullOrWhiteSpace (text)) {
				return BuildFailure (string.Empty);
			}
			return Execute (new KernelCaseCreationRequest {
				CustomerName = customerName,
				Mode = KernelCaseCreationMode.CreateCaseBatch,
				SelectedFolderPath = text
			}, kernelWorkbook, expectedSystemRoot);
		}

		private KernelCaseCreationResult Execute (KernelCaseCreationRequest request, Workbook kernelWorkbook, string expectedSystemRoot)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			CreatedCasePresentationWaitService.WaitSession waitSession = null;
			long waitUiShownElapsedMs = -1L;
			bool waitSessionTransferred = false;
			try {
				ValidateRequest (request);
				ValidateBoundKernelWorkbookOrThrow (kernelWorkbook, expectedSystemRoot, "Execute");
				_logger.Info ("Kernel case command validated. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				if (ShouldShowCreatedCaseWaitUi (request.Mode)) {
					waitSession = _createdCasePresentationWaitService.ShowWaiting (stopwatch);
					waitSession.UpdateStage (CreatedCasePresentationWaitService.CreatingStageTitle);
					if (request.Mode == KernelCaseCreationMode.NewCaseDefault) {
						waitUiShownElapsedMs = stopwatch.ElapsedMilliseconds;
					}
				}
				KernelCaseCreationResult kernelCaseCreationResult = _kernelCaseCreationService.CreateCase (kernelWorkbook, expectedSystemRoot, request);
				if (!kernelCaseCreationResult.Success) {
					return kernelCaseCreationResult;
				}
				if (request.Mode == KernelCaseCreationMode.NewCaseDefault && waitUiShownElapsedMs >= 0L) {
					_logger.Info ("NewCaseDefault timing. segment=waitUiShownToCaseCreated, caseWorkbookPath=" + (kernelCaseCreationResult.CaseWorkbookPath ?? string.Empty) + ", elapsedMs=" + Math.Max (0L, stopwatch.ElapsedMilliseconds - waitUiShownElapsedMs));
				}
				if (!ShouldPromptToOpenCreatedCase (kernelCaseCreationResult.Mode)) {
					if (kernelCaseCreationResult.Mode == KernelCaseCreationMode.CreateCaseBatch) {
						waitSession?.UpdateStage (CreatedCasePresentationWaitService.BatchOpeningFolderStageTitle);
					}
					PresentCaseFolderBestEffort (kernelCaseCreationResult, "KernelCaseCreationCommandService.Execute.NoPrompt");
					if (kernelCaseCreationResult.Mode == KernelCaseCreationMode.CreateCaseBatch) {
						_logger.Info ("NewCaseDefault timing detail. segment=batchStage, phase=stage3Update, mode=CreateCaseBatch, method=Execute, threadId=" + Environment.CurrentManagedThreadId + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
						waitSession?.UpdateStage (CreatedCasePresentationWaitService.BatchReturningHomeStageTitle);
					}
					kernelCaseCreationResult.ShouldCloseKernelHome = false;
					if (kernelCaseCreationResult.Mode == KernelCaseCreationMode.CreateCaseBatch) {
						CloseWaitSession (waitSession, restoreOwner: true);
						waitSession = null;
					}
					_logger.Info ("Kernel case command completed without open prompt. mode=" + request.Mode.ToString () + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
					return kernelCaseCreationResult;
				}
				waitSessionTransferred = true;
				KernelCaseCreationResult kernelCaseCreationResult2 = CompleteInteractiveOpenFlow (kernelCaseCreationResult, stopwatch, waitSession);
				waitSession = null;
				return kernelCaseCreationResult2;
			} catch (Exception exception) {
				_logger.Error ("Kernel case creation failed.", exception);
				return BuildFailure (CreateCaseFailedMessage);
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
				if (ShouldOfferCreatedCaseFolderOnClose (kernelCaseCreationResult.Mode)) {
					_caseWorkbookLifecycleService.MarkCreatedCaseFolderOfferPending (kernelCaseCreationResult.CreatedWorkbook);
				}
				kernelCaseCreationResult.ShouldCloseKernelHome = true;
				_logger.Info ("Kernel case presentation complete. mode=" + kernelCaseCreationResult.Mode.ToString () + ", success=" + kernelCaseCreationResult.Success + ", presentationOutcome=" + kernelCaseCreationResult.PresentationOutcome.ToString () + ", presentationOutcomeReason=" + SanitizeForSingleLine (kernelCaseCreationResult.PresentationOutcomeReason) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds);
				return kernelCaseCreationResult;
			} catch (Exception exception) {
				_logger.Error ("Kernel case open after save failed. mode=" + result.Mode.ToString () + ", presentationOutcome=" + CasePresentationOutcome.Failed.ToString () + ", presentationOutcomeReason=" + ResolvePresentationFailureReason (result, exception), exception);
				return BuildOpenFailure (result, OpenCaseFailedMessage, ResolvePresentationFailureReason (result, exception));
			}
		}

		private static bool ShouldPromptToOpenCreatedCase (KernelCaseCreationMode mode)
		{
			return mode == KernelCaseCreationMode.NewCaseDefault || mode == KernelCaseCreationMode.CreateCaseSingle;
		}

		private static bool ShouldOfferCreatedCaseFolderOnClose (KernelCaseCreationMode mode)
		{
			return mode == KernelCaseCreationMode.NewCaseDefault || mode == KernelCaseCreationMode.CreateCaseSingle;
		}

		private static bool ShouldShowCreatedCaseWaitUi (KernelCaseCreationMode mode)
		{
			return ShouldPromptToOpenCreatedCase (mode) || mode == KernelCaseCreationMode.CreateCaseBatch;
		}

		private void PresentCaseFolderBestEffort (KernelCaseCreationResult result, string reason)
		{
			if (!ShouldPresentCaseFolder (result)) {
				return;
			}
			_kernelCasePresentationService.OpenCaseFolderAndWait (result.CaseFolderPath, reason);
		}

		private static bool ShouldPresentCaseFolder (KernelCaseCreationResult result)
		{
			return result != null && result.Success && !string.IsNullOrWhiteSpace (result.CaseFolderPath);
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

		private string SelectFolderAndRemember (Workbook workbook, string expectedSystemRoot)
		{
			Stopwatch stopwatch = Stopwatch.StartNew ();
			if (!TryValidateBoundKernelWorkbook (workbook, expectedSystemRoot, "SelectFolderAndRemember")) {
				return string.Empty;
			}
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

		private bool TryValidateBoundKernelWorkbook (Workbook workbook, string expectedSystemRoot, string operationName)
		{
			if (workbook == null) {
				_logger.Warn ("Kernel case command failed closed because bound workbook was null. operation=" + (operationName ?? string.Empty));
				return false;
			}
			if (!_kernelWorkbookService.IsKernelWorkbook (workbook)) {
				_logger.Warn ("Kernel case command failed closed because workbook was not kernel. operation=" + (operationName ?? string.Empty) + ", workbook=" + GetWorkbookIdentity (workbook));
				return false;
			}
			string text = NormalizeSystemRoot (_kernelCasePathService.ResolveSystemRoot (workbook));
			string text2 = NormalizeSystemRoot (expectedSystemRoot);
			if (string.IsNullOrWhiteSpace (text2) || string.IsNullOrWhiteSpace (text) || !string.Equals (text, text2, StringComparison.OrdinalIgnoreCase)) {
				_logger.Warn ("Kernel case command failed closed because kernel root mismatched. operation=" + (operationName ?? string.Empty) + ", workbook=" + GetWorkbookIdentity (workbook) + ", actualSystemRoot=" + text + ", expectedSystemRoot=" + text2);
				return false;
			}
			return true;
		}

		private void ValidateBoundKernelWorkbookOrThrow (Workbook workbook, string expectedSystemRoot, string operationName)
		{
			if (!TryValidateBoundKernelWorkbook (workbook, expectedSystemRoot, operationName)) {
				throw new InvalidOperationException ("Bound Kernel workbook was not available.");
			}
		}

		private string GetWorkbookIdentity (Workbook workbook)
		{
			return _excelInteropService.GetWorkbookFullName (workbook);
		}

		private static string NormalizeSystemRoot (string systemRoot)
		{
			string text = (systemRoot ?? string.Empty).Trim ();
			if (string.IsNullOrWhiteSpace (text)) {
				return string.Empty;
			}
			text = text.TrimEnd (Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
			try {
				return Path.GetFullPath (text).TrimEnd (Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
			} catch {
				return text;
			}
		}

		private static KernelCaseCreationResult BuildFailure (string userMessage)
		{
			return new KernelCaseCreationResult {
				Success = false,
				UserMessage = (userMessage ?? string.Empty),
				ShouldCloseKernelHome = false,
				PresentationOutcome = CasePresentationOutcome.NotStarted,
				PresentationOutcomeReason = string.Empty
			};
		}

		private static KernelCaseCreationResult BuildOpenFailure (KernelCaseCreationResult source, string userMessage, string presentationOutcomeReason)
		{
			return new KernelCaseCreationResult {
				Success = false,
				Mode = (source == null) ? KernelCaseCreationMode.NewCaseDefault : source.Mode,
				CaseFolderPath = (source == null) ? string.Empty : source.CaseFolderPath,
				CaseWorkbookPath = (source == null) ? string.Empty : source.CaseWorkbookPath,
				CreatedWorkbook = (source == null) ? null : source.CreatedWorkbook,
				UserMessage = (userMessage ?? string.Empty),
				ShouldCloseKernelHome = false,
				PresentationOutcome = CasePresentationOutcome.Failed,
				PresentationOutcomeReason = presentationOutcomeReason ?? string.Empty
			};
		}

		private static string ResolvePresentationFailureReason (KernelCaseCreationResult result, Exception exception)
		{
			if (result != null && !string.IsNullOrWhiteSpace (result.PresentationOutcomeReason)) {
				return SanitizeForSingleLine (result.PresentationOutcomeReason);
			}
			return "OpenCreatedCaseException:" + ((exception == null) ? string.Empty : exception.GetType ().Name);
		}

		private static string SanitizeForSingleLine (string value)
		{
			return (value ?? string.Empty).Replace ("\r\n", " | ").Replace ("\n", " | ").Replace ("\r", " | ");
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
