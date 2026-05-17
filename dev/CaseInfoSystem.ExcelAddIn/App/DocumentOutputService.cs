using System;
using System.Diagnostics;
using System.IO;
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
			_logger.Debug ("DocumentOutputService.BuildDocumentOutputPath", "Completed elapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " " + BuildFolderDiagnostics ("folder", text) + ", " + BuildValueDiagnostics ("baseName", text2) + ", " + BuildPathDiagnostics ("outputPath", text3));
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
				_logger.Info ("DocumentOutputService resolved workbook folder. " + BuildPathDiagnostics ("source", text3) + ", " + BuildFolderDiagnostics ("resolved", text5));
				return text5;
			}
			return string.Empty;
		}

		private static string BuildValueDiagnostics (string label, string value)
		{
			string safeLabel = label ?? string.Empty;
			string safeValue = value ?? string.Empty;
			return safeLabel + "Provided=" + (!string.IsNullOrWhiteSpace (safeValue))
				+ ", " + safeLabel + "Length=" + safeValue.Length;
		}

		private static string BuildPathDiagnostics (string label, string path)
		{
			string safeLabel = label ?? string.Empty;
			string safePath = path ?? string.Empty;
			return safeLabel + "Present=" + (!string.IsNullOrWhiteSpace (safePath))
				+ ", " + safeLabel + "Length=" + safePath.Length
				+ ", " + safeLabel + "Extension=" + SafeGetExtension (safePath)
				+ ", " + safeLabel + "Exists=" + SafeFileExists (safePath);
		}

		private static string BuildFolderDiagnostics (string label, string path)
		{
			string safeLabel = label ?? string.Empty;
			string safePath = path ?? string.Empty;
			return safeLabel + "Present=" + (!string.IsNullOrWhiteSpace (safePath))
				+ ", " + safeLabel + "Length=" + safePath.Length
				+ ", " + safeLabel + "Exists=" + SafeDirectoryExists (safePath);
		}

		private static string SafeGetExtension (string path)
		{
			try {
				return Path.GetExtension (path ?? string.Empty) ?? string.Empty;
			} catch {
				return string.Empty;
			}
		}

		private static bool SafeFileExists (string path)
		{
			if (string.IsNullOrWhiteSpace (path)) {
				return false;
			}
			try {
				return File.Exists (path);
			} catch {
				return false;
			}
		}

		private static bool SafeDirectoryExists (string path)
		{
			if (string.IsNullOrWhiteSpace (path)) {
				return false;
			}
			try {
				return Directory.Exists (path);
			} catch {
				return false;
			}
		}

		private static string FormatElapsedSeconds (TimeSpan elapsed)
		{
			return elapsed.TotalSeconds.ToString ("0.000");
		}
	}
}
