using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class KernelWorkbookResolverService
	{
		private readonly Application _application;

		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		internal KernelWorkbookResolverService (Application application, ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
		}

		internal Workbook ResolveOrOpen (Workbook caseWorkbook, out bool openedNow)
		{
			return ResolveOrOpenCore (caseWorkbook, out openedNow, readOnly: false);
		}

		internal Workbook ResolveOrOpenReadOnly (Workbook caseWorkbook, out bool openedNow)
		{
			return ResolveOrOpenCore (caseWorkbook, out openedNow, readOnly: true);
		}

		private Workbook ResolveOrOpenCore (Workbook caseWorkbook, out bool openedNow, bool readOnly)
		{
			openedNow = false;
			if (caseWorkbook == null) {
				return null;
			}
			string text = ResolveKernelPath (caseWorkbook);
			if (string.IsNullOrWhiteSpace (text)) {
				return null;
			}
			Workbook workbook = _excelInteropService.FindOpenWorkbook (text);
			if (workbook != null) {
				return workbook;
			}
			if (!_pathCompatibilityService.FileExistsSafe (text)) {
				return null;
			}
			bool enableEvents = _application.EnableEvents;
			try {
				_application.EnableEvents = false;
				Workbook workbook2 = _application.Workbooks.Open (text, 0, readOnly, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing);
				HideWorkbookWindows (workbook2);
				openedNow = true;
				return workbook2;
			} finally {
				_application.EnableEvents = enableEvents;
			}
		}

		private string ResolveKernelPath (Workbook caseWorkbook)
		{
			string text = _pathCompatibilityService.NormalizePath (_excelInteropService.TryGetDocumentProperty (caseWorkbook, "SYSTEM_ROOT"));
			if (string.IsNullOrWhiteSpace (text)) {
				return string.Empty;
			}
			try {
				return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath (text, _pathCompatibilityService);
			} catch {
				return string.Empty;
			}
		}

		private static void HideWorkbookWindows (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			Windows windows = null;
			try {
				windows = workbook.Windows;
				int windowCount = (windows == null) ? 0 : windows.Count;
				for (int index = 1; index <= windowCount; index++) {
					Window window = null;
					try {
						window = windows [index];
						if (window != null) {
							window.Visible = false;
						}
					} finally {
						if (window != null && Marshal.IsComObject (window)) {
							ComObjectReleaseService.Release (window);
						}
					}
				}
			} catch {
			} finally {
				if (windows != null && Marshal.IsComObject (windows)) {
					ComObjectReleaseService.Release (windows);
				}
			}
		}
	}
}
