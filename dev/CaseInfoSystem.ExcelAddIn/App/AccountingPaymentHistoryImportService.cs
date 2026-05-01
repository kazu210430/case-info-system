using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingPaymentHistoryImportService
	{
		private const double VbaDpi = 96.0;

		private const double VbaPointsPerInch = 72.0;

		private const string PromptAnchorCellAddress = "AA1";

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly UserErrorService _userErrorService;

		private readonly Logger _logger;

		private AccountingImportRangePromptForm _activePromptForm;

		private Workbook _activePromptWorkbook;

		private string _activePromptWorkbookFullName;

		internal AccountingPaymentHistoryImportService (AccountingWorkbookService accountingWorkbookService, UserErrorService userErrorService, Logger logger)
		{
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_userErrorService = userErrorService ?? throw new ArgumentNullException ("userErrorService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (WorkbookContext context)
		{
			if (context == null) {
				throw new ArgumentNullException ("context");
			}
			Workbook workbook = context.Workbook;
			if (workbook == null) {
				throw new InvalidOperationException ("会計書類セットブックが見つかりません。");
			}
			Worksheet worksheet = _accountingWorkbookService.GetWorksheet (workbook, "会計依頼書");
			Worksheet worksheet2 = _accountingWorkbookService.GetWorksheet (workbook, "お支払い履歴");
			ValidatePaymentHistoryExists (worksheet2);
			string selectedTargetAddress = TryResolveTargetCellAddress (context, workbook, worksheet);
			int startRound = ReadRoundValue (worksheet2, "A13", 1);
			int endRound = ResolveLatestRound (worksheet2);
			_accountingWorkbookService.HighlightAccountingImportTargets (workbook);
			ShowPrompt (context, workbook, worksheet, worksheet2, selectedTargetAddress, startRound, endRound);
		}

		internal void HandleWorkbookBeforeClose (Workbook workbook)
		{
			if (IsSameWorkbook (_activePromptWorkbook, workbook)) {
				CloseActivePrompt (clearHighlight: true);
			}
		}

		internal void HandleSheetActivated (Workbook workbook, string activeSheetCodeName)
		{
			if (IsSameWorkbook (_activePromptWorkbook, workbook) && !string.Equals (activeSheetCodeName, "会計依頼書", StringComparison.OrdinalIgnoreCase)) {
				_logger.Info ("Accounting payment history import prompt closed on sheet activation. activeSheet=" + (activeSheetCodeName ?? string.Empty));
				CloseActivePrompt (clearHighlight: true);
			}
		}

		private void ApplyImport (Workbook workbook, Worksheet requestWorksheet, Worksheet paymentHistoryWorksheet, Range targetCell, AccountingImportRange importRange)
		{
			int num = ResolvePaymentHistoryBaseRowOffset (paymentHistoryWorksheet);
			int num2 = importRange.StartRound + num;
			int num3 = importRange.EndRound + num;
			if (num3 < num2) {
				throw new InvalidOperationException ("終期は始期以上で指定してください。");
			}
			double num4 = SumColumn (paymentHistoryWorksheet, num2, num3, "D");
			double num5 = SumColumn (paymentHistoryWorksheet, num2, num3, "G");
			double num6 = SumColumn (paymentHistoryWorksheet, num2, num3, "H");
			string text = SafeAddress (targetCell);
			_logger.Info ("Accounting payment history import start. startRound=" + importRange.StartRound + ", endRound=" + importRange.EndRound + ", startRow=" + num2 + ", endRow=" + num3 + ", targetAddress=" + text);
			try {
				_accountingWorkbookService.WriteCellValue (requestWorksheet, text, num4);
				_accountingWorkbookService.WriteCellValue (workbook, "会計依頼書", "F24", num5);
				_accountingWorkbookService.WriteCellValue (workbook, "会計依頼書", "F25", num6);
				_accountingWorkbookService.WriteCell (workbook, "会計依頼書", "B6", "以下のとおり会計処理をお願いします。\r\n\r\n別紙「お支払い履歴」の第" + importRange.StartRound + "回から第" + importRange.EndRound + "回の合計です。");
				if (num5 > 0.0) {
					_accountingWorkbookService.WriteCell (workbook, "会計依頼書", "J24", "各回の源泉税の合計です。");
				}
				_logger.Info ("Accounting payment history import completed. amountTotal=" + num4 + ", taxTotal=" + num5 + ", expenseTotal=" + num6 + ", targetAddress=" + SafeAddress (targetCell));
			} catch (Exception exception) {
				_logger.Error ("Accounting payment history import failed during write phase.", exception);
				throw;
			}
		}

		private void ValidatePaymentHistoryExists (Worksheet paymentHistoryWorksheet)
		{
			if (paymentHistoryWorksheet == null) {
				throw new InvalidOperationException ("お支払い履歴シートが見つかりません。");
			}
			object value = _accountingWorkbookService.ReadCellValue (paymentHistoryWorksheet, "B13");
			if (string.IsNullOrWhiteSpace (Convert.ToString (value))) {
				throw new InvalidOperationException ("お支払い履歴を先に作成してください。");
			}
		}

		private int ResolvePaymentHistoryBaseRowOffset (Worksheet paymentHistoryWorksheet)
		{
			int num = ReadRoundValue (paymentHistoryWorksheet, "A13", 1);
			return (num == 1) ? 12 : 13;
		}

		private int ResolveLatestRound (Worksheet paymentHistoryWorksheet)
		{
			Range range = null;
			try {
				range = ((dynamic)paymentHistoryWorksheet.Cells [12, "B"]).End [XlDirection.xlDown] as Range;
				int num = range?.Row ?? 13;
				int num2 = ReadRoundValue (paymentHistoryWorksheet, "A13", 1);
				return (num2 == 1) ? (num - 12) : (num - 13);
			} finally {
				ReleaseComObject (range);
			}
		}

		private int ReadRoundValue (Worksheet worksheet, string address, int defaultValue)
		{
			try {
				object value = _accountingWorkbookService.ReadCellValue (worksheet, address);
				int result;
				return int.TryParse (Convert.ToString (value), out result) ? result : defaultValue;
			} catch {
				return defaultValue;
			}
		}

		private Range ResolveTargetCell (WorkbookContext context, Workbook workbook, Worksheet requestWorksheet, string preservedAddress)
		{
			string a = SafeActiveSheetCodeName (context.Workbook);
			if (!string.Equals (a, "会計依頼書", StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("シートを会計依頼書に切り替えてください。");
			}
			Range range = TryGetPreservedTargetCell (requestWorksheet, preservedAddress);
			if (range != null) {
				return range;
			}
			Range range2 = null;
			Range range3 = null;
			Range range4 = null;
			try {
				range2 = TryGetActiveCell (workbook);
				range3 = ((_Worksheet)requestWorksheet).get_Range ((object)"F15:F20", Type.Missing);
				range4 = ((range2 == null) ? null : requestWorksheet.Application.Intersect (range2, range3, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));
				if (range4 == null) {
					throw new InvalidOperationException ("金額を入力したいセルを1つ選択してください（黄色エリア内）。");
				}
				return range2;
			} finally {
				ReleaseComObject (range4);
				ReleaseComObject (range3);
			}
		}

		private string TryResolveTargetCellAddress (WorkbookContext context, Workbook workbook, Worksheet requestWorksheet)
		{
			Range range = null;
			Range range2 = null;
			Range range3 = null;
			try {
				string a = SafeActiveSheetCodeName (context.Workbook);
				if (!string.Equals (a, "会計依頼書", StringComparison.OrdinalIgnoreCase)) {
					return string.Empty;
				}
				range = TryGetActiveCell (workbook);
				if (range == null) {
					return string.Empty;
				}
				range2 = ((_Worksheet)requestWorksheet).get_Range ((object)"F15:F20", Type.Missing);
				range3 = requestWorksheet.Application.Intersect (range, range2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				return (range3 == null) ? string.Empty : SafeAddress (range);
			} catch {
				return string.Empty;
			} finally {
				ReleaseComObject (range3);
				ReleaseComObject (range2);
			}
		}

		private static Range TryGetActiveCell (Workbook workbook)
		{
			try {
				Range range = (workbook?.Application)?.ActiveCell;
				if (range == null) {
					return null;
				}
				return (range.Cells.Count == 1) ? range : null;
			} catch {
				return null;
			}
		}

		private void ShowPrompt (WorkbookContext context, Workbook workbook, Worksheet requestWorksheet, Worksheet paymentHistoryWorksheet, string selectedTargetAddress, int startRound, int endRound)
		{
			CloseActivePrompt (clearHighlight: true);
			ExcelWindowOwner owner = ExcelWindowOwner.From (context.Window);
			AccountingImportRangePromptForm form = new AccountingImportRangePromptForm (startRound, endRound);
			ApplySheetAnchoredLocation (form, context.Window, requestWorksheet, PromptAnchorCellAddress);
			_activePromptForm = form;
			_activePromptWorkbook = workbook;
			_activePromptWorkbookFullName = SafeWorkbookFullName (workbook);
			form.Confirmed += delegate(object sender, AccountingImportRangePromptConfirmedEventArgs e) {
				try {
					Range targetCell = ResolveTargetCell (context, workbook, requestWorksheet, selectedTargetAddress);
					ApplyImport (workbook, requestWorksheet, paymentHistoryWorksheet, targetCell, e.ImportRange);
					form.CloseByCode ();
				} catch (Exception exception) {
					_logger.Error ("Accounting payment history import prompt confirmed handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingControl_ActionInvoked", exception);
				}
			};
			form.Canceled += delegate {
				_logger.Info ("Accounting payment history import prompt closed.");
				if (workbook != null) {
					try {
						_accountingWorkbookService.ClearAccountingImportTargetHighlight (workbook);
					} catch (Exception exception) {
						_logger.Error ("Accounting payment history import prompt highlight cleanup failed.", exception);
					}
				}
				if (_activePromptForm == form) {
					_activePromptForm = null;
					_activePromptWorkbook = null;
					_activePromptWorkbookFullName = string.Empty;
				}
				if (owner != null) {
					owner.Dispose ();
				}
			};
			form.ShowModeless (owner);
			_logger.Info ("Accounting payment history import prompt shown.");
		}

		private void ApplySheetAnchoredLocation (Form form, Window window, Worksheet worksheet, string anchorAddress)
		{
			if (form != null) {
				System.Drawing.Point? point = TryCalculateSheetAnchoredLocation (window, worksheet, anchorAddress);
				if (point.HasValue) {
					form.StartPosition = FormStartPosition.Manual;
					form.Location = point.Value;
				}
			}
		}

		private System.Drawing.Point? TryCalculateSheetAnchoredLocation (Window window, Worksheet worksheet, string anchorAddress)
		{
			Range range = null;
			try {
				if (window == null || worksheet == null || string.IsNullOrWhiteSpace (anchorAddress)) {
					return null;
				}
				range = ((_Worksheet)worksheet).get_Range ((object)anchorAddress, Type.Missing);
				double num = Convert.ToDouble ((dynamic)window.Zoom);
				double a = Convert.ToDouble ((dynamic)range.Left) * num / 100.0 * 96.0 / 72.0;
				double a2 = Convert.ToDouble ((dynamic)range.Top) * num / 100.0 * 96.0 / 72.0;
				return new System.Drawing.Point (window.PointsToScreenPixelsX (0) + Convert.ToInt32 (Math.Round (a)), window.PointsToScreenPixelsY (0) + Convert.ToInt32 (Math.Round (a2)));
			} catch (Exception exception) {
				_logger.Error ("Accounting payment history import prompt location calculation failed.", exception);
				return null;
			} finally {
				ReleaseComObject (range);
			}
		}

		private void CloseActivePrompt (bool clearHighlight)
		{
			if (_activePromptForm == null) {
				return;
			}
			try {
				if (clearHighlight && _activePromptWorkbook != null) {
					_accountingWorkbookService.ClearAccountingImportTargetHighlight (_activePromptWorkbook);
				}
				if (!_activePromptForm.IsDisposed) {
					_activePromptForm.Close ();
				}
			} catch {
			} finally {
				_activePromptForm = null;
				_activePromptWorkbook = null;
				_activePromptWorkbookFullName = string.Empty;
			}
		}

		private static string SafeWorkbookFullName (Workbook workbook)
		{
			try {
				return (workbook == null) ? string.Empty : (workbook.FullName ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private static bool IsSameWorkbook (Workbook left, Workbook right)
		{
			if (left == null || right == null) {
				return false;
			}
			if (left == right) {
				return true;
			}
			try {
				return string.Equals (left.FullName ?? string.Empty, right.FullName ?? string.Empty, StringComparison.OrdinalIgnoreCase);
			} catch {
				return false;
			}
		}

		private static Range TryGetPreservedTargetCell (Worksheet requestWorksheet, string preservedAddress)
		{
			if (requestWorksheet == null || string.IsNullOrWhiteSpace (preservedAddress)) {
				return null;
			}
			Range range = null;
			Range range2 = null;
			Range range3 = null;
			try {
				range = ((_Worksheet)requestWorksheet).get_Range ((object)preservedAddress, Type.Missing);
				range2 = ((_Worksheet)requestWorksheet).get_Range ((object)"F15:F20", Type.Missing);
				range3 = requestWorksheet.Application.Intersect (range, range2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				if (range3 == null) {
					ReleaseComObject (range);
					return null;
				}
				return range;
			} catch {
				ReleaseComObject (range);
				return null;
			} finally {
				ReleaseComObject (range3);
				ReleaseComObject (range2);
			}
		}

		private static double SumColumn (Worksheet worksheet, int startRow, int endRow, string columnName)
		{
			Range range = null;
			try {
				range = worksheet.Range [(dynamic)worksheet.Cells [startRow, columnName], (dynamic)worksheet.Cells [endRow, columnName]];
				object value = worksheet.Application.WorksheetFunction.Sum (range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				return Convert.ToDouble (value);
			} finally {
				ReleaseComObject (range);
			}
		}

		private static string SafeActiveSheetCodeName (Workbook workbook)
		{
			try {
				Worksheet worksheet = ((workbook == null) ? null : (workbook.ActiveSheet as Worksheet));
				return (worksheet == null) ? string.Empty : (worksheet.CodeName ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private static string SafeAddress (Range range)
		{
			try {
				return (range == null) ? string.Empty : (Convert.ToString (range.get_Address ((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing)) ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private static void ReleaseComObject (object comObject)
		{
			CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (comObject);
		}
	}
}
