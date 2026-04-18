using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class NavigationService
	{
		private readonly IExcelInteropService _excelInteropService;

		private readonly IWorkbookRoleResolver _workbookRoleResolver;

		private readonly Logger _logger;

		internal NavigationService (IExcelInteropService excelInteropService, IWorkbookRoleResolver workbookRoleResolver, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal WorkbookContext ResolveActiveContext ()
		{
			Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook ();
			Window activeWindow = _excelInteropService.GetActiveWindow ();
			return CreateContext (activeWorkbook, activeWindow);
		}

		internal WorkbookContext CreateContext (Workbook workbook, Window window)
		{
			if (workbook == null) {
				return null;
			}
			WorkbookRole role = _workbookRoleResolver.Resolve (workbook);
			string systemRoot = _workbookRoleResolver.ResolveSystemRoot (workbook);
			string workbookFullName = _excelInteropService.GetWorkbookFullName (workbook);
			string activeSheetCodeName = _excelInteropService.GetActiveSheetCodeName (workbook);
			Window window2 = ResolvePreferredWindow (workbook, window);
			return CreateContextFromSeed (new WorkbookContextSeed (workbook, window2, role, systemRoot, workbookFullName, activeSheetCodeName));
		}

		internal WorkbookContext CreateContextFromSeed (WorkbookContextSeed seed)
		{
			if (seed == null) {
				return null;
			}
			return new WorkbookContext (seed.Workbook, seed.Window, seed.Role, seed.SystemRoot, seed.WorkbookFullName, seed.ActiveSheetCodeName);
		}

		private string ResolveActiveSheetIdentity (Workbook workbook, WorkbookRole role)
		{
			if (role != WorkbookRole.Accounting) {
				return _excelInteropService.GetActiveSheetCodeName (workbook);
			}
			try {
				Worksheet worksheet = ((workbook == null) ? null : (workbook.ActiveSheet as Worksheet));
				return (worksheet == null) ? string.Empty : (worksheet.CodeName ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private Window ResolvePreferredWindow (Workbook workbook, Window window)
		{
			if (window != null) {
				return window;
			}
			Window activeWindow = _excelInteropService.GetActiveWindow ();
			if (activeWindow != null) {
				Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook ();
				string workbookFullName = _excelInteropService.GetWorkbookFullName (activeWorkbook);
				string workbookFullName2 = _excelInteropService.GetWorkbookFullName (workbook);
				if (string.Equals (workbookFullName, workbookFullName2, StringComparison.OrdinalIgnoreCase)) {
					return activeWindow;
				}
			}
			return _excelInteropService.GetFirstVisibleWindow (workbook);
		}

		internal bool TryNavigateToWorkbook (Workbook workbook, string reason)
		{
			bool result = _excelInteropService.ActivateWorkbook (workbook);
			_logger.Info ("TryNavigateToWorkbook result=" + result + ", reason=" + (reason ?? string.Empty));
			return result;
		}

		internal bool TryNavigateToSheet (Workbook workbook, string sheetCodeName, string reason)
		{
			bool result = _excelInteropService.ActivateWorkbook (workbook) && _excelInteropService.ActivateWorksheetByCodeName (workbook, sheetCodeName);
			_logger.Info ("TryNavigateToSheet result=" + result + ", reason=" + (reason ?? string.Empty) + ", sheetCodeName=" + (sheetCodeName ?? string.Empty));
			return result;
		}

		internal void TraceContext (WorkbookContext context, string reason)
		{
			if (context == null) {
				_logger.Warn ("TraceContext skipped because context was null. reason=" + (reason ?? string.Empty));
				return;
			}
			_logger.Info ("WorkbookContext resolved. reason=" + (reason ?? string.Empty) + ", role=" + context.Role.ToString () + ", workbook=" + context.WorkbookFullName + ", systemRoot=" + context.SystemRoot + ", activeSheetCodeName=" + context.ActiveSheetCodeName);
		}
	}
}
