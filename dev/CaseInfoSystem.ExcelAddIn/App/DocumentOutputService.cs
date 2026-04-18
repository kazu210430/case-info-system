using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class DocumentOutputService
	{
		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly Logger _logger;

		internal DocumentOutputService (ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal string BuildOutputFileName (Workbook workbook, string documentName, string customerName)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			string text = _excelInteropService.TryGetDocumentProperty (workbook, "NAME_RULE_A");
			string text2 = _excelInteropService.TryGetDocumentProperty (workbook, "NAME_RULE_B");
			return KernelNamingService.BuildDocumentName (string.IsNullOrWhiteSpace (text) ? "YYYY" : text, string.IsNullOrWhiteSpace (text2) ? "DOC" : text2, (documentName ?? string.Empty).Trim (), (customerName ?? string.Empty).Trim (), DateTime.Today);
		}

		internal string BuildDocumentOutputPath (Workbook workbook, string documentName, string customerName)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			Stopwatch stopwatch = Stopwatch.StartNew ();
			string text = ResolveWorkbookFolder (workbook);
			if (string.IsNullOrWhiteSpace (text)) {
				_logger.Info ("DocumentOutputService could not resolve workbook folder. elapsed=" + FormatElapsedSeconds (stopwatch.Elapsed));
				return string.Empty;
			}
			string text2 = BuildOutputFileName (workbook, documentName, customerName);
			string text3 = _pathCompatibilityService.BuildUniquePath (text, text2, ".docx");
			_logger.Debug ("DocumentOutputService.BuildDocumentOutputPath", "Completed elapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " folder=" + text + " baseName=" + text2 + " outputPath=" + text3);
			return text3;
		}

		internal string ResolveWorkbookOutputExtension (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			try {
				if (workbook.FileFormat == XlFileFormat.xlOpenXMLWorkbook) {
					return ".xlsx";
				}
			} catch {
			}
			string workbookFullName = _excelInteropService.GetWorkbookFullName (workbook);
			return WorkbookFileNameResolver.GetWorkbookExtensionOrDefault (workbookFullName);
		}

		internal string PrepareSavePath (string rawFullPath)
		{
			if (string.IsNullOrWhiteSpace (rawFullPath)) {
				return string.Empty;
			}
			string text = _pathCompatibilityService.BuildSafeSavePath (rawFullPath);
			return (text.Length == 0) ? string.Empty : _pathCompatibilityService.EnsureUniquePathStandard (text);
		}

		internal string ResolveWorkbookFolder (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			string text = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookPath (workbook));
			if (text.Length > 0) {
				string text2 = _pathCompatibilityService.ResolveToExistingLocalPath (text);
				if (text2.Length > 0) {
					return text2;
				}
			}
			string text3 = _pathCompatibilityService.NormalizePath (_excelInteropService.GetWorkbookFullName (workbook));
			string text4 = _pathCompatibilityService.GetParentFolderPath (text3);
			if (text4.Length == 0) {
				text4 = text3;
			}
			string text5 = _pathCompatibilityService.ResolveToExistingLocalPath (text4);
			if (text5.Length > 0) {
				_logger.Info ("DocumentOutputService resolved workbook folder. source=" + text3 + " resolved=" + text5);
				return text5;
			}
			return string.Empty;
		}

		private static string FormatElapsedSeconds (TimeSpan elapsed)
		{
			return elapsed.TotalSeconds.ToString ("0.000");
		}
	}
}
