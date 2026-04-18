using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class AccountingTemplateResolver
	{
		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly Logger _logger;

		internal AccountingTemplateResolver (ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal string ResolveTemplatePath (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			string text = ResolveSystemRoot (workbook);
			if (string.IsNullOrWhiteSpace (text)) {
				throw new InvalidOperationException ("SYSTEM_ROOT を取得できませんでした。");
			}
			string text2 = _pathCompatibilityService.CombinePath (text, "雛形");
			_logger.Info ("Accounting template resolve start. workbook=" + SafeWorkbookName (workbook) + ", systemRoot=" + text + ", templateDirectory=" + text2);
			if (!Directory.Exists (text2)) {
				_logger.Warn ("Accounting template folder was not found: " + text2);
				throw new InvalidOperationException ("雛形フォルダが見つかりません: " + text2);
			}
			string[] array = ResolveTemplateCandidates (text2);
			_logger.Info ("Accounting template candidates count=" + array.Length);
			if (array.Length == 0) {
				_logger.Warn ("Accounting template candidate was not found in directory: " + text2);
				throw new InvalidOperationException ("雛形フォルダに会計書類セットの Excel ブックが見つかりません: " + text2);
			}
			if (array.Length > 1) {
				_logger.Warn ("Accounting template candidates were duplicated in directory: " + text2);
				throw new InvalidOperationException ("雛形フォルダに会計書類セットの Excel ブックが複数あります: " + text2);
			}
			string text3 = _pathCompatibilityService.ResolveToExistingLocalPath (array [0]);
			if (string.IsNullOrWhiteSpace (text3)) {
				text3 = _pathCompatibilityService.NormalizePath (array [0]);
			}
			_logger.Info ("Accounting template resolved. path=" + text3);
			return text3;
		}

		private static string[] ResolveTemplateCandidates (string templateDirectory)
		{
			List<string> list = new List<string> ();
			string[] supportedMainWorkbookExtensions = WorkbookFileNameResolver.GetSupportedMainWorkbookExtensions ();
			for (int i = 0; i < supportedMainWorkbookExtensions.Length; i++) {
				string searchPattern = "会計書類セット*" + supportedMainWorkbookExtensions [i];
				string[] files = Directory.GetFiles (templateDirectory, searchPattern, SearchOption.TopDirectoryOnly);
				if (files.Length != 0) {
					list.AddRange (files);
				}
			}
			return list.ToArray ();
		}

		private static string SafeWorkbookName (Workbook workbook)
		{
			try {
				return (workbook == null) ? string.Empty : (workbook.FullName ?? workbook.Name ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private string ResolveSystemRoot (Workbook workbook)
		{
			string text = _pathCompatibilityService.NormalizePath (_excelInteropService.TryGetDocumentProperty (workbook, "SYSTEM_ROOT"));
			if (!string.IsNullOrWhiteSpace (text)) {
				return text;
			}
			string text2 = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookPath (workbook));
			if (!string.IsNullOrWhiteSpace (text2)) {
				return text2;
			}
			string fullPath = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookFullName (workbook));
			return _pathCompatibilityService.GetParentFolderPath (fullPath);
		}
	}
}
