using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class DocumentSaveService
	{
		private readonly DocumentOutputService _documentOutputService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly LocalWorkCopyService _localWorkCopyService;

		private readonly WordInteropService _wordInteropService;

		private readonly Logger _logger;

		private readonly DocumentSaveServiceTestHooks _testHooks;

		internal DocumentSaveService (DocumentOutputService documentOutputService, PathCompatibilityService pathCompatibilityService, LocalWorkCopyService localWorkCopyService, WordInteropService wordInteropService, Logger logger)
			: this (documentOutputService, pathCompatibilityService, localWorkCopyService, wordInteropService, logger, testHooks: null)
		{
		}

		internal DocumentSaveService (DocumentOutputService documentOutputService, PathCompatibilityService pathCompatibilityService, LocalWorkCopyService localWorkCopyService, WordInteropService wordInteropService, Logger logger, DocumentSaveServiceTestHooks testHooks)
		{
			_documentOutputService = documentOutputService ?? throw new ArgumentNullException ("documentOutputService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_localWorkCopyService = localWorkCopyService ?? throw new ArgumentNullException ("localWorkCopyService");
			_wordInteropService = wordInteropService ?? throw new ArgumentNullException ("wordInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_testHooks = testHooks;
		}

		internal DocumentSaveResult SaveDocument (object wordApplication, object wordDocument, string requestedFinalPath)
		{
			if (wordDocument == null) {
				throw new ArgumentNullException ("wordDocument");
			}
			string text = PrepareSavePath (requestedFinalPath);
			if (text.Length == 0) {
				throw new InvalidOperationException ("保存先パスを準備できませんでした。");
			}
			if (!IsUnderSyncRoot (text)) {
				string savedPath = SaveDocumentAsDocx (wordDocument, text);
				return new DocumentSaveResult (savedPath, text, isLocalWorkCopy: false);
			}
			string text2 = BuildLocalWorkCopyPath (text);
			if (text2.Length == 0) {
				_logger.Info ("DocumentSaveService fell back to direct save because local work copy path was not prepared. final=" + text);
				string savedPath2 = SaveDocumentAsDocx (wordDocument, text);
				return new DocumentSaveResult (savedPath2, text, isLocalWorkCopy: false);
			}
			string text3 = SaveDocumentAsDocx (wordDocument, text2);
			RegisterLocalWorkCopy (wordApplication, text3, text);
			_logger.Info ("DocumentSaveService saved via local work copy. local=" + text3 + ", final=" + text);
			return new DocumentSaveResult (text3, text, isLocalWorkCopy: true);
		}

		private string PrepareSavePath (string requestedFinalPath)
		{
			if (_testHooks != null && _testHooks.PrepareSavePath != null) {
				return _testHooks.PrepareSavePath (requestedFinalPath) ?? string.Empty;
			}
			return _documentOutputService.PrepareSavePath (requestedFinalPath);
		}

		private bool IsUnderSyncRoot (string finalPath)
		{
			return (_testHooks != null && _testHooks.IsUnderSyncRoot != null) ? _testHooks.IsUnderSyncRoot (finalPath) : _pathCompatibilityService.IsUnderSyncRoot (finalPath);
		}

		private string BuildLocalWorkCopyPath (string finalPath)
		{
			if (_testHooks != null && _testHooks.BuildLocalWorkCopyPath != null) {
				return _testHooks.BuildLocalWorkCopyPath (finalPath) ?? string.Empty;
			}
			return _localWorkCopyService.BuildLocalWorkCopyPath (finalPath);
		}

		private string SaveDocumentAsDocx (object wordDocument, string savePath)
		{
			if (_testHooks != null && _testHooks.SaveDocumentAsDocx != null) {
				return _testHooks.SaveDocumentAsDocx (wordDocument, savePath) ?? string.Empty;
			}
			return _wordInteropService.SaveDocumentAsDocx (wordDocument, savePath);
		}

		private void RegisterLocalWorkCopy (object wordApplication, string localPath, string finalPath)
		{
			if (_testHooks != null && _testHooks.RegisterLocalWorkCopy != null) {
				_testHooks.RegisterLocalWorkCopy (wordApplication, localPath, finalPath);
				return;
			}
			_localWorkCopyService.RegisterLocalWorkCopy (wordApplication, localPath, finalPath);
		}

		internal sealed class DocumentSaveServiceTestHooks
		{
			internal Func<string, string> PrepareSavePath { get; set; }

			internal Func<string, bool> IsUnderSyncRoot { get; set; }

			internal Func<string, string> BuildLocalWorkCopyPath { get; set; }

			internal Func<object, string, string> SaveDocumentAsDocx { get; set; }

			internal Action<object, string, string> RegisterLocalWorkCopy { get; set; }
		}
	}
}
