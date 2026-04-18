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

		internal DocumentSaveService (DocumentOutputService documentOutputService, PathCompatibilityService pathCompatibilityService, LocalWorkCopyService localWorkCopyService, WordInteropService wordInteropService, Logger logger)
		{
			_documentOutputService = documentOutputService ?? throw new ArgumentNullException ("documentOutputService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_localWorkCopyService = localWorkCopyService ?? throw new ArgumentNullException ("localWorkCopyService");
			_wordInteropService = wordInteropService ?? throw new ArgumentNullException ("wordInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal DocumentSaveResult SaveDocument (object wordApplication, object wordDocument, string requestedFinalPath)
		{
			if (wordDocument == null) {
				throw new ArgumentNullException ("wordDocument");
			}
			string text = _documentOutputService.PrepareSavePath (requestedFinalPath);
			if (text.Length == 0) {
				throw new InvalidOperationException ("保存先パスを準備できませんでした。");
			}
			if (!_pathCompatibilityService.IsUnderSyncRoot (text)) {
				string savedPath = _wordInteropService.SaveDocumentAsDocx (wordDocument, text);
				return new DocumentSaveResult (savedPath, text, isLocalWorkCopy: false);
			}
			string text2 = _localWorkCopyService.BuildLocalWorkCopyPath (text);
			if (text2.Length == 0) {
				_logger.Info ("DocumentSaveService fell back to direct save because local work copy path was not prepared. final=" + text);
				string savedPath2 = _wordInteropService.SaveDocumentAsDocx (wordDocument, text);
				return new DocumentSaveResult (savedPath2, text, isLocalWorkCopy: false);
			}
			string text3 = _wordInteropService.SaveDocumentAsDocx (wordDocument, text2);
			_localWorkCopyService.RegisterLocalWorkCopy (wordApplication, text3, text);
			_logger.Info ("DocumentSaveService saved via local work copy. local=" + text3 + ", final=" + text);
			return new DocumentSaveResult (text3, text, isLocalWorkCopy: true);
		}
	}
}
