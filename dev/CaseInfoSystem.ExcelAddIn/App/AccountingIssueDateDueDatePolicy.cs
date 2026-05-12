namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingIssueDateDueDatePolicy
	{
		internal const string NoPaymentNoticeText = "（お支払い頂く金額はございません）";

		internal static bool ShouldWriteNoPaymentNotice (double invoicePaymentAmount)
		{
			return invoicePaymentAmount == 0.0;
		}
	}
}
