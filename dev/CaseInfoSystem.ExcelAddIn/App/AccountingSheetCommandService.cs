using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingSheetCommandService
	{
		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly Logger _logger;

		internal AccountingSheetCommandService (AccountingWorkbookService accountingWorkbookService, Logger logger)
		{
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (WorkbookContext context, string actionId)
		{
			if (context == null) {
				throw new ArgumentNullException ("context");
			}
			if (!string.IsNullOrWhiteSpace (actionId)) {
				switch (actionId) {
				case "import-estimate-to-current":
					CopyBetweenForms (context.Workbook, "見積書", SafeActiveSheetName (context));
					break;
				case "import-invoice-to-current":
					CopyBetweenForms (context.Workbook, "請求書", SafeActiveSheetName (context));
					break;
				case "import-receipt-to-current":
					CopyBetweenForms (context.Workbook, "領収書", SafeActiveSheetName (context));
					break;
				case "reset-current-sheet":
					ResetSheet (context.Workbook, SafeActiveSheetName (context));
					break;
				}
			}
		}

		private void CopyBetweenForms (Workbook workbook, string sourceSheetName, string targetSheetName)
		{
			if (string.IsNullOrWhiteSpace (targetSheetName)) {
				throw new InvalidOperationException ("対象シートを特定できません。");
			}
			if (string.Equals (targetSheetName, "会計依頼書", StringComparison.OrdinalIgnoreCase)) {
				_accountingWorkbookService.CopyFormulaRange (workbook, sourceSheetName, "A4:X4", targetSheetName, "A4:X4");
				_accountingWorkbookService.CopyFormulaRange (workbook, sourceSheetName, "A14:Y44", targetSheetName, "A14:Y44");
				_accountingWorkbookService.CopyFormulaRange (workbook, sourceSheetName, "A3", targetSheetName, "AB3");
				_accountingWorkbookService.CopyValueRange (workbook, targetSheetName, "Y3", targetSheetName, "A3");
				_logger.Info ("Accounting request sheet imported from source. source=" + sourceSheetName + ", target=" + targetSheetName);
			} else {
				_accountingWorkbookService.CopyFormulaRange (workbook, sourceSheetName, "A3:X4", targetSheetName, "A3:X4");
				_accountingWorkbookService.CopyFormulaRange (workbook, sourceSheetName, "A14:Y44", targetSheetName, "A14:Y44");
				_logger.Info ("Accounting form sheet imported from source. source=" + sourceSheetName + ", target=" + targetSheetName);
			}
		}

		private void ResetSheet (Workbook workbook, string targetSheetName)
		{
			string text = BuildResetMessage (targetSheetName);
			DialogResult dialogResult = MessageBox.Show (text, "案件情報System", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
			if (dialogResult != DialogResult.OK) {
				_logger.Info ("Accounting sheet reset canceled. sheet=" + targetSheetName);
				return;
			}
			switch (targetSheetName) {
			case "見積書":
				ResetEstimateLikeSheet (workbook, targetSheetName, resetComment: true, clearDueDate: false);
				break;
			case "請求書":
				ResetEstimateLikeSheet (workbook, targetSheetName, resetComment: false, clearDueDate: true);
				break;
			case "領収書":
				ResetEstimateLikeSheet (workbook, targetSheetName, resetComment: false, clearDueDate: false);
				break;
			case "会計依頼書":
				ResetAccountingRequestSheet (workbook);
				break;
			default:
				throw new InvalidOperationException ("未対応のリセット対象です: " + targetSheetName);
			}
		}

		private void ResetEstimateLikeSheet (Workbook workbook, string targetSheetName, bool resetComment, bool clearDueDate)
		{
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "Y3:Z3", targetSheetName, "Y3:Z3");
			if (resetComment) {
				_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "A8", targetSheetName, "A8");
			}
			ApplyCommonRecovery (workbook, targetSheetName);
			if (clearDueDate) {
				_accountingWorkbookService.ClearMergeAreaContents (workbook, targetSheetName, "G10");
			}
			_logger.Info ("Accounting form sheet reset completed. sheet=" + targetSheetName);
		}

		private void ResetAccountingRequestSheet (Workbook workbook)
		{
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "Y4:Z4", "会計依頼書", "Y3:Z3");
			ApplyCommonRecovery (workbook, "会計依頼書");
			_accountingWorkbookService.WriteCell (workbook, "会計依頼書", "B6", "以下のとおり会計処理をお願いします。");
			_logger.Info ("Accounting request sheet reset completed.");
		}

		private void ApplyCommonRecovery (Workbook workbook, string targetSheetName)
		{
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "A1:Y1", targetSheetName, "A1:Y1");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "G12", targetSheetName, "G12");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "G13", targetSheetName, "G13");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "B15:Y16", targetSheetName, "B15:Y16");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "B17:B22", targetSheetName, "B17:B22");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "C27:C31", targetSheetName, "C27:C31");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "F21:F31", targetSheetName, "F21:F31");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "J24:Y24", targetSheetName, "J24:Y24");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "J25:J31", targetSheetName, "J25:J31");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "B33", targetSheetName, "B33");
			_accountingWorkbookService.CopyFormulaRange (workbook, "リカバリ", "J17:J20", targetSheetName, "J17:J20");
			_accountingWorkbookService.WriteCell (workbook, targetSheetName, "A3", "依頼者");
			_accountingWorkbookService.WriteCell (workbook, targetSheetName, "A4", "件名");
			_accountingWorkbookService.ClearMergeAreaContents (workbook, targetSheetName, "F33");
			_accountingWorkbookService.ClearMergeAreaContents (workbook, targetSheetName, "J33");
			_accountingWorkbookService.ClearMergeAreaContents (workbook, targetSheetName, "B35");
			_accountingWorkbookService.WriteCell (workbook, targetSheetName, "F17:F20", string.Empty);
			_accountingWorkbookService.WriteCell (workbook, targetSheetName, "K17:K20", string.Empty);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F15", 0);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F16", 0);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F17", 0);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F18", 0);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F19", 0);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F20", 0);
			_accountingWorkbookService.SetInteriorColorIndex (workbook, targetSheetName, "F33", 0);
		}

		private static string SafeActiveSheetName (WorkbookContext context)
		{
			return (context == null) ? string.Empty : (context.ActiveSheetCodeName ?? string.Empty);
		}

		private static string BuildResetMessage (string targetSheetName)
		{
			switch (targetSheetName) {
			case "見積書":
				return "見積書は全てクリアされます。よろしいですか？";
			case "請求書":
				return "請求書は全てクリアされます。よろしいですか？";
			case "領収書":
				return "領収書は全てクリアされます。よろしいですか？";
			case "会計依頼書":
				return "会計依頼書は全てクリアされます。よろしいですか？";
			default:
				return "対象シートは全てクリアされます。よろしいですか？";
			}
		}
	}
}
