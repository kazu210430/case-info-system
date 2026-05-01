using System;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingSetKernelSyncService
	{
		private sealed class UserDataTransferSnapshot
		{
			internal string PostalCode { get; }

			internal string Address1 { get; }

			internal string Address2 { get; }

			internal string PhoneNumber { get; }

			internal string NameLine1 { get; }

			internal string NameLine2 { get; }

			internal string SourceKernelPath { get; }

			internal UserDataTransferSnapshot (string postalCode, string address1, string address2, string phoneNumber, string nameLine1, string nameLine2, string sourceKernelPath)
			{
				PostalCode = postalCode ?? string.Empty;
				Address1 = address1 ?? string.Empty;
				Address2 = address2 ?? string.Empty;
				PhoneNumber = phoneNumber ?? string.Empty;
				NameLine1 = nameLine1 ?? string.Empty;
				NameLine2 = nameLine2 ?? string.Empty;
				SourceKernelPath = sourceKernelPath ?? string.Empty;
			}
		}

		private sealed class AccountingSetTransferPlan
		{
			internal string SourceKernelPath { get; }

			internal string AddressLine { get; }

			internal string AddressLineWithBreak { get; }

			internal string NameLine1 { get; }

			internal string NameLine2 { get; }

			internal AccountingSetTransferPlan (string sourceKernelPath, string addressLine, string addressLineWithBreak, string nameLine1, string nameLine2)
			{
				SourceKernelPath = sourceKernelPath ?? string.Empty;
				AddressLine = addressLine ?? string.Empty;
				AddressLineWithBreak = addressLineWithBreak ?? string.Empty;
				NameLine1 = nameLine1 ?? string.Empty;
				NameLine2 = nameLine2 ?? string.Empty;
			}
		}

		private readonly ExcelInteropService _excelInteropService;

		private readonly AccountingTemplateResolver _accountingTemplateResolver;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly Logger _logger;

		internal AccountingSetKernelSyncService (ExcelInteropService excelInteropService, AccountingTemplateResolver accountingTemplateResolver, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_accountingTemplateResolver = accountingTemplateResolver ?? throw new ArgumentNullException ("accountingTemplateResolver");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (Workbook kernelWorkbook)
		{
			if (kernelWorkbook == null) {
				throw new ArgumentNullException ("kernelWorkbook");
			}
			Worksheet worksheet = ResolveUserDataWorksheet (kernelWorkbook);
			if (worksheet == null) {
				throw new InvalidOperationException ("利用者情報シートが見つかりません。");
			}
			string text = _accountingTemplateResolver.ResolveTemplatePath (kernelWorkbook);
			string text2 = _pathCompatibilityService.ResolveToExistingLocalPath (text);
			if (string.IsNullOrWhiteSpace (text2)) {
				text2 = text;
			}
			Workbook workbook = _excelInteropService.FindOpenWorkbook (text) ?? _excelInteropService.FindOpenWorkbook (text2);
			bool flag = workbook != null;
			Application application = null;
			string workbookFullName = _excelInteropService.GetWorkbookFullName (kernelWorkbook);
			_logger.Info ("Accounting set kernel sync start. kernelWorkbook=" + workbookFullName + ", template=" + text + ", localTemplate=" + text2 + ", alreadyOpen=" + flag);
			try {
				if (!flag) {
					Application obj = (Application)Activator.CreateInstance (Marshal.GetTypeFromCLSID (new Guid ("00024500-0000-0000-C000-000000000046")));
					obj.Visible = false;
					obj.DisplayAlerts = false;
					obj.ScreenUpdating = false;
					obj.EnableEvents = false;
					application = obj;
					_logger.Info ("Accounting set kernel sync hidden Excel session created. appHwnd=" + SafeApplicationHwnd (application));
					workbook = application.Workbooks.Open (text2, 0, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
					_logger.Debug ("AccountingSetKernelSyncService", "Template opened in hidden Excel application.");
				}
				if (workbook == null) {
					throw new InvalidOperationException ("会計書類セットを開けませんでした: " + text2);
				}
				UserDataTransferSnapshot snapshot = ReadTransferSnapshot (worksheet, _excelInteropService.GetWorkbookFullName (kernelWorkbook));
				AccountingSetTransferPlan transferPlan = BuildTransferPlan (snapshot);
				ApplyTransferPlan (workbook, transferPlan);
				workbook.Save ();
				_logger.Info ("Accounting set kernel sync completed. path=" + text2 + ", alreadyOpen=" + flag);
			} catch (Exception exception) {
				_logger.Error ("Accounting set kernel sync failed. kernelWorkbook=" + workbookFullName + ", template=" + text + ", localTemplate=" + text2 + ", alreadyOpen=" + flag, exception);
				throw;
			} finally {
				if (!flag) {
					CloseWorkbookQuietly (workbook);
					_logger.Info ("Accounting set kernel sync hidden Excel session quitting. appHwnd=" + SafeApplicationHwnd (application));
					QuitApplicationQuietly (application);
				}
				ReleaseComObject (worksheet);
			}
		}

		private static UserDataTransferSnapshot ReadTransferSnapshot (Worksheet userDataWorksheet, string kernelWorkbookFullName)
		{
			return new UserDataTransferSnapshot (ReadUserDataValue (userDataWorksheet, 0), ReadUserDataValue (userDataWorksheet, 1), ReadUserDataValue (userDataWorksheet, 2), ReadUserDataValue (userDataWorksheet, 3), ReadUserDataValue (userDataWorksheet, 6), ReadUserDataValue (userDataWorksheet, 7), kernelWorkbookFullName);
		}

		private static AccountingSetTransferPlan BuildTransferPlan (UserDataTransferSnapshot snapshot)
		{
			if (snapshot == null) {
				throw new ArgumentNullException ("snapshot");
			}
			string addressLine = BuildUserAddressLine (snapshot.PostalCode, snapshot.Address1, snapshot.Address2, snapshot.PhoneNumber, withLineBreak: false);
			string addressLineWithBreak = BuildUserAddressLine (snapshot.PostalCode, snapshot.Address1, snapshot.Address2, snapshot.PhoneNumber, withLineBreak: true);
			return new AccountingSetTransferPlan (snapshot.SourceKernelPath, addressLine, addressLineWithBreak, snapshot.NameLine1, snapshot.NameLine2);
		}

		private void ApplyTransferPlan (Workbook accountingWorkbook, AccountingSetTransferPlan transferPlan)
		{
			if (accountingWorkbook == null) {
				throw new ArgumentNullException ("accountingWorkbook");
			}
			if (transferPlan == null) {
				throw new ArgumentNullException ("transferPlan");
			}
			_logger.Info ("Accounting set kernel transfer values. addressLine=" + transferPlan.AddressLine + ", addressLineWithBreak=" + transferPlan.AddressLineWithBreak.Replace ("\n", " / ") + ", nameLine1=" + transferPlan.NameLine1 + ", nameLine2=" + transferPlan.NameLine2);
			_excelInteropService.SetDocumentProperty (accountingWorkbook, "SOURCE_KERNEL_PATH", transferPlan.SourceKernelPath);
			string[] sheetNames = new string[3] { "見積書", "請求書", "領収書" };
			string[] sheetNames2 = new string[2] { "分割払い予定表", "お支払い履歴" };
			_accountingWorkbookService.WriteSameValueToSheets (accountingWorkbook, sheetNames, "A40", transferPlan.AddressLine);
			_accountingWorkbookService.WriteSameValueToSheets (accountingWorkbook, sheetNames2, "A5", transferPlan.AddressLineWithBreak);
			_accountingWorkbookService.WriteCell (accountingWorkbook, "請求書", "G7", transferPlan.NameLine1);
			_accountingWorkbookService.WriteCell (accountingWorkbook, "請求書", "G8", transferPlan.NameLine2);
			_accountingWorkbookService.WriteCell (accountingWorkbook, "分割払い予定表", "A7", transferPlan.NameLine1);
			_accountingWorkbookService.WriteCell (accountingWorkbook, "分割払い予定表", "A8", transferPlan.NameLine2);
			_accountingWorkbookService.WriteCell (accountingWorkbook, "分割払い予定表", "A5", transferPlan.AddressLineWithBreak);
			_logger.Debug ("AccountingSetKernelSyncService", "Kernel transfer cell writes completed.");
		}

		private Worksheet ResolveUserDataWorksheet (Workbook kernelWorkbook)
		{
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (kernelWorkbook, "shUserData");
			if (worksheet != null) {
				return worksheet;
			}
			try {
				return kernelWorkbook.Worksheets ["ユーザー情報"] as Worksheet;
			} catch {
				return null;
			}
		}

		private static string BuildUserAddressLine (string postalCode, string address1, string address2, string phoneNumber, bool withLineBreak)
		{
			string text = (postalCode ?? string.Empty).Trim ();
			string text2 = (address1 ?? string.Empty).Trim ();
			string text3 = (address2 ?? string.Empty).Trim ();
			string text4 = (phoneNumber ?? string.Empty).Trim ();
			if (withLineBreak) {
				return "〒" + text + "\u3000" + text2 + "\n" + new string ('\u3000', 11) + text3 + "\u3000℡" + text4;
			}
			return "〒" + text + "\u3000" + text2 + "\u3000" + text3 + "\u3000℡" + text4;
		}

		private static string ReadUserDataValue (Worksheet userDataWorksheet, int rowOffset)
		{
			if (userDataWorksheet == null) {
				return string.Empty;
			}
			try {
				object obj = ((dynamic)userDataWorksheet.Cells [2 + rowOffset, "B"]).Value2;
				return (Convert.ToString (obj) ?? string.Empty).Trim ();
			} catch {
				return string.Empty;
			}
		}

		private static void CloseWorkbookQuietly (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			try {
				workbook.Close (false, Type.Missing, Type.Missing);
			} catch {
			} finally {
				ReleaseComObject (workbook);
			}
		}

		private static void QuitApplicationQuietly (Application application)
		{
			if (application == null) {
				return;
			}
			try {
				application.Quit ();
			} catch {
			} finally {
				ReleaseComObject (application);
			}
		}

		private static void ReleaseComObject (object comObject)
		{
			// 同期用に所有した非表示 Excel 参照は完全解放の方針を維持する。
			CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease (comObject);
		}

		private static string SafeApplicationHwnd (Application application)
		{
			try {
				return (application == null) ? string.Empty : (Convert.ToString (application.Hwnd) ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}
	}
}
