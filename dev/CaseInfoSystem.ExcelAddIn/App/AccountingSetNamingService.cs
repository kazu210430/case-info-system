using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingSetNamingService
	{
		private readonly DocumentOutputService _documentOutputService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		internal AccountingSetNamingService (DocumentOutputService documentOutputService, PathCompatibilityService pathCompatibilityService)
		{
			_documentOutputService = documentOutputService ?? throw new ArgumentNullException ("documentOutputService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
		}

		internal string BuildCaseOutputPath (Workbook workbook, string outputFolderPath, string customerName, string templatePath)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (string.IsNullOrWhiteSpace (outputFolderPath)) {
				throw new InvalidOperationException ("会計書類セットの保存先フォルダを解決できませんでした。");
			}
			string text = _documentOutputService.BuildOutputFileName (workbook, "会計書類セット", customerName);
			if (string.IsNullOrWhiteSpace (text)) {
				text = "会計書類セット";
			}
			string workbookExtensionOrDefault = WorkbookFileNameResolver.GetWorkbookExtensionOrDefault (templatePath);
			return _pathCompatibilityService.BuildUniquePath (outputFolderPath, text, workbookExtensionOrDefault);
		}
	}
}
