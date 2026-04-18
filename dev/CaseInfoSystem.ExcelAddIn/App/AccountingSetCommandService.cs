using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingSetCommandService
	{
		private readonly WorkbookRoleResolver _workbookRoleResolver;

		private readonly AccountingSetCreateService _accountingSetCreateService;

		private readonly AccountingSetKernelSyncService _accountingSetKernelSyncService;

		private readonly Logger _logger;

		internal AccountingSetCommandService (WorkbookRoleResolver workbookRoleResolver, AccountingSetCreateService accountingSetCreateService, AccountingSetKernelSyncService accountingSetKernelSyncService, Logger logger)
		{
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_accountingSetCreateService = accountingSetCreateService ?? throw new ArgumentNullException ("accountingSetCreateService");
			_accountingSetKernelSyncService = accountingSetKernelSyncService ?? throw new ArgumentNullException ("accountingSetKernelSyncService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (Workbook workbook)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			WorkbookRole workbookRole = _workbookRoleResolver.Resolve (workbook);
			_logger.Info ("Accounting set command invoked. workbook=" + SafeWorkbookName (workbook) + ", role=" + workbookRole);
			switch (workbookRole) {
			case WorkbookRole.Case:
				_logger.Debug ("AccountingSetCommandService", "Routing to CASE create flow.");
				_accountingSetCreateService.Execute (workbook);
				break;
			case WorkbookRole.Kernel:
				_logger.Debug ("AccountingSetCommandService", "Routing to Kernel sync flow.");
				_accountingSetKernelSyncService.Execute (workbook);
				break;
			default:
				_logger.Warn ("AccountingSetCommandService rejected workbook because role was unsupported.");
				throw new InvalidOperationException ("会計書類セットは CASE または Kernel ブックからのみ実行できます。");
			}
		}

		private static string SafeWorkbookName (Workbook workbook)
		{
			try {
				return (workbook == null) ? string.Empty : (workbook.FullName ?? workbook.Name ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}
	}
}
