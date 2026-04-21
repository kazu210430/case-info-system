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
				throw new InvalidOperationException ("Save path could not be resolved.");
			}

			DocumentSaveOutcome saveOutcome = SaveViaAdjacentTempReplace (wordApplication, wordDocument, finalPath);
			_logger.Info ("DocumentSaveService save completed. final=" + finalPath + ", totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed));
			return new DocumentSaveResult (saveOutcome.SavedPath, finalPath, isLocalWorkCopy: false, saveOutcome.ActiveDocument);
		}

		private string PrepareSavePath (string requestedFinalPath)
		{
			if (_testHooks != null && _testHooks.PrepareSavePath != null) {
				return _testHooks.PrepareSavePath (requestedFinalPath) ?? string.Empty;
			}
			return _documentOutputService.PrepareSavePath (requestedFinalPath);
		}

		private string SaveDocumentCopyAsDocx (object wordDocument, string savePath)
		{
			if (_testHooks != null && _testHooks.SaveDocumentCopyAsDocx != null) {
				return _testHooks.SaveDocumentCopyAsDocx (wordDocument, savePath) ?? string.Empty;
			}
			return _wordInteropService.SaveDocumentCopyAsDocx (wordDocument, savePath);
		}

		private object OpenDocument (object wordApplication, string fullPath)
		{
			if (_testHooks != null && _testHooks.OpenDocument != null) {
				return _testHooks.OpenDocument (wordApplication, fullPath);
			}
			return _wordInteropService.OpenDocument (wordApplication, fullPath);
		}

		private DocumentSaveOutcome SaveViaAdjacentTempReplace (object wordApplication, object wordDocument, string finalPath)
		{
			Stopwatch totalStopwatch = Stopwatch.StartNew ();
			Stopwatch phaseStopwatch = Stopwatch.StartNew ();
			string stagingPath = BuildStagingPath (finalPath);
			bool finalExists = FileExistsSafe (finalPath);
			bool isUnderSyncRoot = _pathCompatibilityService.IsUnderSyncRoot (finalPath);
			_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "SaveContext mode=" + FormatSaveMode (finalExists) + " location=" + FormatSaveLocation (isUnderSyncRoot) + " elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
			try {
				string stagingSavedPath = SaveDocumentCopyAsDocx (wordDocument, stagingPath);
				if (string.IsNullOrWhiteSpace (stagingSavedPath) || !FileExistsSafe (stagingSavedPath)) {
					throw new IOException ("Staging save failed.");
				}
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "StagingSaved elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " staging=" + stagingSavedPath);
				phaseStopwatch.Restart ();
				if (!_pathCompatibilityService.PromoteAdjacentStagingFileSafe (stagingSavedPath, finalPath)) {
					throw new IOException ("Atomic replace failed.");
				}
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "FinalReplaced elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
				phaseStopwatch.Restart ();
				object activeDocument = OpenDocument (wordApplication, finalPath);
				if (activeDocument == null) {
					throw new IOException ("Reopen after save failed.");
				}
				_logger.Debug ("DocumentSaveService.SaveDirectWithBackup", "FinalReopened elapsed=" + FormatElapsedSeconds (phaseStopwatch.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (totalStopwatch.Elapsed) + " final=" + finalPath);
				object originalDocument = wordDocument;
				_wordInteropService.CloseDocumentNoSave (ref originalDocument);
				return new DocumentSaveOutcome (finalPath, activeDocument);
			} catch {
				TryDeleteFileQuietly (stagingPath);
				throw;
			}
		}

		private bool FileExistsSafe (string path)
		{
			return _pathCompatibilityService.FileExistsSafe (path);
		}

		private static string BuildStagingPath (string finalPath)
		{
			string directoryName = Path.GetDirectoryName (finalPath) ?? string.Empty;
			string fileNameWithoutExtension = Path.GetFileNameWithoutExtension (finalPath);
			string extension = Path.GetExtension (finalPath);
			if (fileNameWithoutExtension.Length == 0) {
				fileNameWithoutExtension = "document";
			}
			string fileName = fileNameWithoutExtension + ".tmp_" + Guid.NewGuid ().ToString ("N") + extension;
			return Path.Combine (directoryName, fileName);
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

			internal Func<object, string, string> SaveDocumentCopyAsDocx { get; set; }

			internal Func<object, string, object> OpenDocument { get; set; }
		}

		private sealed class DocumentSaveOutcome
		{
			internal DocumentSaveOutcome (string savedPath, object activeDocument)
			{
				SavedPath = savedPath ?? string.Empty;
				ActiveDocument = activeDocument;
			}

			internal string SavedPath { get; }

			internal object ActiveDocument { get; }
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
