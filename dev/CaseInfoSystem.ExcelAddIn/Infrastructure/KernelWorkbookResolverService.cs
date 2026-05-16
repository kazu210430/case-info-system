using System;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class KernelWorkbookAccessResult
	{
		private readonly bool _readOnly;

		private readonly Logger _logger;

		private bool _closeAttempted;

		internal KernelWorkbookAccessResult (string resolvedKernelPath, Workbook workbook, bool workbookWasAlreadyOpen, bool readOnly, Logger logger = null)
		{
			ResolvedKernelPath = resolvedKernelPath ?? string.Empty;
			Workbook = workbook;
			WorkbookWasAlreadyOpen = workbookWasAlreadyOpen;
			_readOnly = readOnly;
			_logger = logger;
		}

		internal string ResolvedKernelPath { get; }

		internal Workbook Workbook { get; }

		internal bool WorkbookWasAlreadyOpen { get; }

		internal bool WorkbookWasOpenedByResolver
		{
			get
			{
				return Workbook != null && !WorkbookWasAlreadyOpen;
			}
		}

		internal void CloseIfOwned ()
		{
			CloseIfOwned (routeName: null, suppressEventsDuringClose: false);
		}

		internal void CloseIfOwned (string routeName)
		{
			CloseIfOwned (routeName, suppressEventsDuringClose: false);
		}

		internal void CloseIfOwned (string routeName, bool suppressEventsDuringClose)
		{
			if (!WorkbookWasOpenedByResolver || _closeAttempted) {
				return;
			}
			_closeAttempted = true;
			if (!suppressEventsDuringClose) {
				CloseOwnedWorkbook (routeName);
				return;
			}
			bool enableEvents = true;
			Application application = null;
			try {
				application = Workbook.Application;
				if (application != null) {
					enableEvents = application.EnableEvents;
					application.EnableEvents = false;
				}
				CloseOwnedWorkbook (routeName);
			} finally {
				if (application != null) {
					application.EnableEvents = enableEvents;
				}
			}
		}

		private void CloseOwnedWorkbook (string routeName)
		{
			string resolvedRouteName = string.IsNullOrWhiteSpace (routeName)
				? nameof (KernelWorkbookResolverService) + "." + nameof (CloseIfOwned)
				: routeName;
			if (_readOnly) {
				WorkbookCloseInteropHelper.CloseReadOnlyWithoutSave (Workbook, _logger, resolvedRouteName);
				return;
			}
			WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave (Workbook, _logger, resolvedRouteName);
		}
	}

	internal sealed class KernelWorkbookResolverService
	{
		private readonly Application _application;

		private readonly ExcelInteropService _excelInteropService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly Logger _logger;

		internal KernelWorkbookResolverService (Application application, ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService, Logger logger = null)
		{
			_application = application ?? throw new ArgumentNullException ("application");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_logger = logger;
		}

		internal KernelWorkbookAccessResult ResolveOrOpen (Workbook caseWorkbook)
		{
			return ResolveOrOpenCore (caseWorkbook, readOnly: false);
		}

		internal KernelWorkbookAccessResult ResolveOrOpenReadOnly (Workbook caseWorkbook)
		{
			return ResolveOrOpenCore (caseWorkbook, readOnly: true);
		}

		private KernelWorkbookAccessResult ResolveOrOpenCore (Workbook caseWorkbook, bool readOnly)
		{
			if (caseWorkbook == null) {
				return CreateResult (string.Empty, null, workbookWasAlreadyOpen: false, readOnly);
			}
			string text = ResolveKernelPath (caseWorkbook);
			if (string.IsNullOrWhiteSpace (text)) {
				return CreateResult (string.Empty, null, workbookWasAlreadyOpen: false, readOnly);
			}
			Workbook workbook = _excelInteropService.FindOpenWorkbook (text);
			if (workbook != null) {
				return CreateResult (text, workbook, workbookWasAlreadyOpen: true, readOnly);
			}
			if (!_pathCompatibilityService.FileExistsSafe (text)) {
				return CreateResult (text, null, workbookWasAlreadyOpen: false, readOnly);
			}
			bool enableEvents = _application.EnableEvents;
			try {
				_application.EnableEvents = false;
				Workbook workbook2 = _application.Workbooks.Open (text, 0, readOnly, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing);
				HideWorkbookWindows (workbook2);
				return CreateResult (text, workbook2, workbookWasAlreadyOpen: false, readOnly);
			} finally {
				_application.EnableEvents = enableEvents;
			}
		}

		private KernelWorkbookAccessResult CreateResult (string resolvedKernelPath, Workbook workbook, bool workbookWasAlreadyOpen, bool readOnly)
		{
			return new KernelWorkbookAccessResult (resolvedKernelPath, workbook, workbookWasAlreadyOpen, readOnly, _logger);
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
						ComObjectReleaseService.Release (window);
					}
				}
			} catch {
			} finally {
				ComObjectReleaseService.Release (windows);
			}
		}
	}
}
