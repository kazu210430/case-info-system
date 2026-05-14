using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
	internal static class AccountingSetSpec
	{
		internal const string WorkbookKindPropertyName = "CASEINFO_WORKBOOK_KIND";

		internal const string WorkbookKindAccountingSetValue = "ACCOUNTING_SET";

		internal const string TemplateFilePrefix = "会計書類セット";

		internal const string TemplateFolderName = "雛形";

		internal const string SourceCasePathPropertyName = "SOURCE_CASE_PATH";

		internal const string SourceKernelPathPropertyName = "SOURCE_KERNEL_PATH";

		internal const string HomeSheetCodeName = "shHOME";

		internal const string UserDataSheetCodeName = "shUserData";

		internal const string UserDataSheetName = "ユーザー情報";

		internal const string InvoiceSheetName = "請求書";

		internal const string EstimateSheetName = "見積書";

		internal const string ReceiptSheetName = "領収書";

		internal const string AccountingRequestSheetName = "会計依頼書";

		internal const string InstallmentSheetName = "分割払い予定表";

		internal const string PaymentHistorySheetName = "お支払い履歴";

		internal const string RecoverySheetName = "リカバリ";

		internal const string ArgumentSheetName = "引数";

		internal const string CustomerWriteCellAddress = "A3";

		internal const string LawyerWriteStartCellAddress = "A41";

		internal const string PaymentHistoryLawyerWriteStartCellAddress = "A6";

		internal const string AccountingAddressCellAddress = "A40";

		internal const string InstallmentAddressCellAddress = "A5";

		internal const string InvoiceNameRow1CellAddress = "G7";

		internal const string InvoiceNameRow2CellAddress = "G8";

		internal const string InstallmentNameRow1CellAddress = "A7";

		internal const string InstallmentNameRow2CellAddress = "A8";

		internal const string AccountingImportTargetRangeAddress = "F15:F20";

		internal const string AccountingImportTaxCellAddress = "F24";

		internal const string AccountingImportExpenseCellAddress = "F25";

		internal const string AccountingImportInstructionCellAddress = "B6";

		internal const string AccountingImportTaxNoteCellAddress = "J24";

		internal const int MaximumLawyerCount = 4;

		internal const int UserDataFirstDataRow = 2;

		internal const int UserDataAccountingNameRow1Offset = 6;

		internal const int UserDataAccountingNameRow2Offset = 7;

		internal static IReadOnlyList<AccountingLawyerReflectionTarget> LawyerReflectionTargets { get; } = new[]
		{
			new AccountingLawyerReflectionTarget (EstimateSheetName, LawyerWriteStartCellAddress),
			new AccountingLawyerReflectionTarget (InvoiceSheetName, LawyerWriteStartCellAddress),
			new AccountingLawyerReflectionTarget (ReceiptSheetName, LawyerWriteStartCellAddress),
			new AccountingLawyerReflectionTarget (AccountingRequestSheetName, LawyerWriteStartCellAddress),
			new AccountingLawyerReflectionTarget (PaymentHistorySheetName, PaymentHistoryLawyerWriteStartCellAddress)
		};
	}

	internal sealed class AccountingLawyerReflectionTarget
	{
		internal AccountingLawyerReflectionTarget (string sheetName, string startCellAddress)
		{
			SheetName = sheetName;
			StartCellAddress = startCellAddress;
		}

		internal string SheetName { get; private set; }

		internal string StartCellAddress { get; private set; }
	}
}
