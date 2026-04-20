using System;
using System.Diagnostics;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class DocumentSaveService
	{
		private readonly DocumentOutputService _documentOutputService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly WordInteropService _wordInteropService;

		private readonly Logger _logger;

		private readonly DocumentSaveServiceTestHooks _testHooks;

		internal DocumentSaveService (DocumentOutputService documentOutputService, WordInteropService wordInteropService, Logger logger)
			: this (documentOutputService, wordInteropService, logger, testHooks: null)
		{
		}

		internal DocumentSaveService (DocumentOutputService documentOutputService, WordInteropService wordInteropService, Logger logger, DocumentSaveServiceTestHooks testHooks)
		{
			_documentOutputService = documentOutputService ?? throw new ArgumentNullException ("documentOutputService");
			_pathCompatibilityService = new PathCompatibilityService ();
			_wordInteropService = wordInteropService ?? throw new ArgumentNullException ("wordInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_testHooks = testHooks;
		}

		internal DocumentSaveResult SaveDocument (object wordApplication, object wordDocument, string requestedFinalPath)
		{
			if (wordDocument == null) {
				throw new ArgumentNullException ("wordDocument");
			}

			Stopwatch totalStopwatch = Stopwatch.StartNew ();
			Stopwatch phaseStopwatch = Stopwatch.StartNew ();
			string finalPath = PrepareSavePath (requestedFinalPath);
			_logger.Debug ("DocumentSaveService.SaveDocument", "PrepareSavePath elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + (finalPath ?? string.Empty));
			if (finalPath.Length == 0) {
				throw new InvalidOperationException ("保存先パスを準備できませんでした。");
			}

			string savedPath = SaveDirectWithBackup (wordDocument, finalPath);
			_logger.Info ("DocumentSaveService direct save completed. final=" + finalPath + ", totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed));
			return new DocumentSaveResult (savedPath, finalPath, isLocalWorkCopy: false);
		}

		private string PrepareSavePath (string requestedFinalPath)
		{
			if (_testHooks != null && _testHooks.PrepareSavePath != null) {
				return _testHooks.PrepareSavePath (requestedFinalPath) ?? string.Empty;
			}
			return _documentOutputService.PrepareSavePath (requestedFinalPath);
		}

		private string SaveDocumentAsDocx (object wordDocument, string savePath)
		{
			if (_testHooks != null && _testHooks.SaveDocumentAsDocx != null) {
				return _testHooks.SaveDocumentAsDocx (wordDocument, savePath) ?? string.Empty;
			}
			return _wordInteropService.SaveDocumentAsDocx (wordDocument, savePath);
		}

		private string SaveDirectWithBackup (object wordDocument, string finalPath)
		{
			Stopwatch totalStopwatch = Stopwatch.StartNew ();
			Stopwatch phaseStopwatch = Stopwatch.StartNew ();
			string backupPath = BuildBackupPath (finalPath);
			bool finalExists = FileExistsSafe (finalPath);
			bool isUnderSyncRoot = _pathCompatibilityService.IsUnderSyncRoot (finalPath);
			_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "SaveContext mode=" + FormatSaveMode (finalExists) + " location=" + FormatSaveLocation (isUnderSyncRoot) + " elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
			bool backupCreated = false;
			if (finalExists) {
				CreateBackupFile (finalPath, backupPath);
				backupCreated = true;
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "BackupCreated elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
			} else {
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "BackupSkipped elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
			}

			phaseStopwatch.Restart ();
			try {
				string savedPath = SaveDocumentAsDocx (wordDocument, finalPath);
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "SaveDocumentAsDocx elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
				phaseStopwatch.Restart ();
				TryDeleteFileQuietly (backupPath);
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "BackupCleanup elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
				return savedPath;
			} catch {
				if (backupCreated) {
					RestoreBackupFileIfPossible (backupPath, finalPath);
					_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "BackupRestore elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
				}
				throw;
			}
		}

		private bool FileExistsSafe (string path)
		{
			return _pathCompatibilityService.FileExistsSafe (path);
		}

		private static string BuildBackupPath (string finalPath)
		{
			string directoryName = Path.GetDirectoryName (finalPath) ?? string.Empty;
			string fileNameWithoutExtension = Path.GetFileNameWithoutExtension (finalPath);
			string extension = Path.GetExtension (finalPath);
			if (fileNameWithoutExtension.Length == 0) {
				fileNameWithoutExtension = "document";
			}
			string fileName = fileNameWithoutExtension + ".bak_" + Guid.NewGuid ().ToString ("N") + extension;
			return Path.Combine (directoryName, fileName);
		}

		private void CreateBackupFile (string finalPath, string backupPath)
		{
			try {
				File.Copy (finalPath, backupPath, overwrite: false);
			} catch (Exception exception) {
				_logger.Error ("DocumentSaveService failed to create backup file before save.", exception);
				throw new IOException ("保存前バックアップの作成に失敗しました。", exception);
			}
		}

		private void RestoreBackupFileIfPossible (string backupPath, string finalPath)
		{
			if (!FileExistsSafe (backupPath)) {
				return;
			}
			try {
				File.Copy (backupPath, finalPath, overwrite: true);
				TryDeleteFileQuietly (backupPath);
			} catch (Exception exception) {
				_logger.Warn ("DocumentSaveService could not restore backup file after save failure. backup=" + backupPath + ", final=" + finalPath + ", error=" + exception.Message);
			}
		}

		private static void TryDeleteFileQuietly (string path)
		{
			if (string.IsNullOrWhiteSpace (path)) {
				return;
			}
			try {
				if (File.Exists (path)) {
					File.Delete (path);
				}
			} catch {
			}
		}

		internal sealed class DocumentSaveServiceTestHooks
		{
			internal Func<string, string> PrepareSavePath { get; set; }

			internal Func<object, string, string> SaveDocumentAsDocx { get; set; }
		}

		private static string FormatElapsedSeconds (TimeSpan elapsed)
		{
			return elapsed.TotalSeconds.ToString ("0.000");
		}

		private static string FormatSaveMode (bool finalExists)
		{
			return finalExists ? "Overwrite" : "CreateNew";
		}

		private static string FormatSaveLocation (bool isUnderSyncRoot)
		{
			return isUnderSyncRoot ? "SyncRoot" : "NonSyncRoot";
		}
	}
}
