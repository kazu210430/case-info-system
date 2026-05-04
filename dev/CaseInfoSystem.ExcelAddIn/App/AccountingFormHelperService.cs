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
	internal sealed class AccountingFormHelperService
	{
		private const double VbaDpi = 96.0;

		private const double VbaPointsPerInch = 72.0;

		private const string SheetAnchorCellAddress = "M1";

		private const string ReverseToolAnchorCellAddress = "AA1";

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly AccountingInstallmentScheduleCommandService _accountingInstallmentScheduleCommandService;

		private readonly AccountingPaymentHistoryCommandService _accountingPaymentHistoryCommandService;

		private readonly AccountingSaveAsService _accountingSaveAsService;

		private readonly UserErrorService _userErrorService;

		private readonly Logger _logger;

		private AccountingReverseGoalSeekForm _activeReverseToolForm;

		private Workbook _activeReverseToolWorkbook;

		private string _activeReverseToolSheetName;

		private AccountingInstallmentScheduleInputForm _activeInstallmentScheduleForm;

		private Workbook _activeInstallmentScheduleWorkbook;

		private ExcelWindowOwner _activeInstallmentScheduleOwner;

		private AccountingPaymentHistoryInputForm _activePaymentHistoryForm;

		private Workbook _activePaymentHistoryWorkbook;

		private ExcelWindowOwner _activePaymentHistoryOwner;

		internal AccountingFormHelperService (AccountingWorkbookService accountingWorkbookService, AccountingInstallmentScheduleCommandService accountingInstallmentScheduleCommandService, AccountingPaymentHistoryCommandService accountingPaymentHistoryCommandService, AccountingSaveAsService accountingSaveAsService, UserErrorService userErrorService, Logger logger)
		{
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_accountingInstallmentScheduleCommandService = accountingInstallmentScheduleCommandService ?? throw new ArgumentNullException ("accountingInstallmentScheduleCommandService");
			_accountingPaymentHistoryCommandService = accountingPaymentHistoryCommandService ?? throw new ArgumentNullException ("accountingPaymentHistoryCommandService");
			_accountingSaveAsService = accountingSaveAsService ?? throw new ArgumentNullException ("accountingSaveAsService");
			_userErrorService = userErrorService ?? throw new ArgumentNullException ("userErrorService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (WorkbookContext context, string actionId)
		{
			if (context == null) {
				throw new ArgumentNullException ("context");
			}
			string text = SafeActiveSheetCodeName (context.Workbook);
			switch (actionId) {
			case "set-issue-date":
				_accountingWorkbookService.CopyValueRange (context.Workbook, text, "Y1", text, "A1");
				break;
			case "set-issue-date-and-due-date":
				ApplyIssueDateAndDueDate (context.Workbook, text);
				break;
			case "open-reverse-tool":
				ShowReverseTool (context, text);
				break;
			case "open-installment-schedule-input":
				EnsureInstallmentScheduleInputVisible (context.Workbook, context.Window, text, activateExisting: true);
				break;
			case "open-payment-history-input":
				ShowPaymentHistoryInput (context, text);
				break;
			}
		}

		internal void HandleSheetActivated (Workbook workbook, Window window, string activeSheetCodeName)
		{
			if (workbook == null || _activeReverseToolWorkbook != workbook || !string.Equals (_activeReverseToolSheetName, activeSheetCodeName, StringComparison.OrdinalIgnoreCase) || !IsMainFormSheet (activeSheetCodeName)) {
				CloseActiveReverseTool (clearHighlight: true);
			}
			if (string.Equals (activeSheetCodeName, "分割払い予定表", StringComparison.OrdinalIgnoreCase)) {
				EnsureInstallmentScheduleInputVisible (workbook, window, activeSheetCodeName, activateExisting: false);
				HidePaymentHistoryInput ();
			} else if (string.Equals (activeSheetCodeName, "お支払い履歴", StringComparison.OrdinalIgnoreCase)) {
				EnsurePaymentHistoryInputVisible (workbook, window, activeSheetCodeName, activateExisting: false);
				HideInstallmentScheduleInput ();
			} else {
				HideInstallmentScheduleInput ();
				HidePaymentHistoryInput ();
			}
		}

		internal void HandleWorkbookBeforeClose (Workbook workbook)
		{
			if (workbook != null) {
				if (IsSameWorkbook (_activeReverseToolWorkbook, workbook)) {
					CloseActiveReverseTool (clearHighlight: true);
				}
				if (IsSameWorkbook (_activeInstallmentScheduleWorkbook, workbook)) {
					CloseActiveInstallmentScheduleInput ();
				}
				if (IsSameWorkbook (_activePaymentHistoryWorkbook, workbook)) {
					CloseActivePaymentHistoryInput ();
				}
			}
		}

		private void ApplyIssueDateAndDueDate (Workbook workbook, string activeSheetCodeName)
		{
			DateTime dateTime = _accountingWorkbookService.ReadDateCell (workbook, activeSheetCodeName, "Y1");
			_accountingWorkbookService.WriteCellValue (workbook, activeSheetCodeName, "A1", dateTime);
			_accountingWorkbookService.WriteCellValue (workbook, activeSheetCodeName, "G10", dateTime.AddDays (14.0));
			_logger.Info ("Accounting issue date and due date applied. sheet=" + activeSheetCodeName + ", issueDate=" + dateTime.ToString ("yyyy-MM-dd"));
		}

		private void ShowReverseTool (WorkbookContext context, string activeSheetCodeName)
		{
			CloseActiveReverseTool (clearHighlight: true);
			_accountingWorkbookService.HighlightReverseToolTargets (context.Workbook, activeSheetCodeName);
			ExcelWindowOwner owner = ExcelWindowOwner.From (context.Window);
			AccountingReverseGoalSeekForm form = new AccountingReverseGoalSeekForm ();
			ApplySheetAnchoredLocation (form, context.Workbook, context.Window, activeSheetCodeName, "AA1");
			_activeReverseToolForm = form;
			_activeReverseToolWorkbook = context.Workbook;
			_activeReverseToolSheetName = activeSheetCodeName;
			form.Confirmed += delegate(object sender, AccountingReverseGoalSeekConfirmedEventArgs e) {
				try {
					ApplyReverseGoalSeek (context.Workbook, activeSheetCodeName, e.Request);
				} catch (Exception exception) {
					_logger.Error ("Accounting reverse tool confirmed handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingReverseGoalSeek", exception);
				}
			};
			form.Canceled += delegate {
				try {
					if (_activeReverseToolWorkbook != null && !string.IsNullOrWhiteSpace (_activeReverseToolSheetName)) {
						_accountingWorkbookService.ClearReverseToolTargets (_activeReverseToolWorkbook, _activeReverseToolSheetName);
					}
				} catch (Exception exception) {
					_logger.Error ("Accounting reverse tool highlight cleanup failed.", exception);
				}
				if (_activeReverseToolForm == form) {
					_activeReverseToolForm = null;
					_activeReverseToolWorkbook = null;
					_activeReverseToolSheetName = string.Empty;
				}
				if (owner != null) {
					owner.Dispose ();
				}
			};
			form.ShowModeless (owner);
			_logger.Info ("Accounting reverse tool shown. sheet=" + activeSheetCodeName);
		}

		private void EnsureInstallmentScheduleInputVisible (Workbook workbook, Window window, string activeSheetCodeName, bool activateExisting)
		{
			if (!string.Equals (activeSheetCodeName, "分割払い予定表", StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("分割払い予定表シートでのみ利用できます。");
			}
			if (workbook == null || window == null) {
				return;
			}
			AccountingInstallmentScheduleFormState state = _accountingInstallmentScheduleCommandService.LoadFormState (workbook);
			if (_activeInstallmentScheduleForm != null && !_activeInstallmentScheduleForm.IsDisposed) {
				ApplySheetAnchoredLocation (_activeInstallmentScheduleForm, workbook, window, "分割払い予定表", "M1");
				_activeInstallmentScheduleForm.BindState (state);
				if (!_activeInstallmentScheduleForm.Visible) {
					_activeInstallmentScheduleForm.Show ();
				}
				if (activateExisting) {
					_activeInstallmentScheduleForm.Activate ();
					_activeInstallmentScheduleForm.FocusInstallmentAmount ();
				}
				return;
			}
			ExcelWindowOwner owner = ExcelWindowOwner.From (window);
			AccountingInstallmentScheduleInputForm form = new AccountingInstallmentScheduleInputForm ();
			ApplySheetAnchoredLocation (form, workbook, window, "分割払い予定表", "M1");
			form.BindState (state);
			_activeInstallmentScheduleForm = form;
			_activeInstallmentScheduleWorkbook = workbook;
			_activeInstallmentScheduleOwner = owner;
			form.IssueDateRequested += delegate {
				try {
					AccountingInstallmentScheduleFormState state2 = _accountingInstallmentScheduleCommandService.ApplyIssueDate (workbook);
					form.BindState (state2);
				} catch (Exception exception) {
					_logger.Error ("Installment schedule issue date handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingInstallmentSchedule.IssueDateRequested", exception);
				}
			};
			form.CreateScheduleRequested += delegate(object sender, AccountingInstallmentScheduleCreateRequestEventArgs e) {
				try {
					AccountingInstallmentScheduleFormState state2 = _accountingInstallmentScheduleCommandService.CreateSchedule (workbook, e.Request);
					form.BindState (state2);
					form.FocusInstallmentAmount ();
				} catch (Exception exception) {
					_logger.Error ("Installment schedule create handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingInstallmentSchedule.CreateScheduleRequested", exception);
				}
			};
			form.ApplyChangeRequested += delegate(object sender, AccountingInstallmentScheduleChangeRequestEventArgs e) {
				try {
					AccountingInstallmentScheduleFormState state2 = _accountingInstallmentScheduleCommandService.ApplyChange (workbook, e.Request);
					form.BindState (state2);
					form.FocusInstallmentAmount ();
				} catch (Exception exception) {
					_logger.Error ("Installment schedule apply change handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingInstallmentSchedule.ApplyChangeRequested", exception);
				}
			};
			form.ResetRequested += delegate {
				try {
					AccountingInstallmentScheduleFormState state2 = _accountingInstallmentScheduleCommandService.Reset (workbook);
					form.BindState (state2);
					form.FocusInstallmentAmount ();
				} catch (Exception exception) {
					_logger.Error ("Installment schedule reset handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingInstallmentSchedule.ResetRequested", exception);
				}
			};
			form.SaveAsRequested += delegate {
				Worksheet worksheet = null;
				try {
					worksheet = workbook.ActiveSheet as Worksheet;
					WorkbookContext context = new WorkbookContext (workbook, window, WorkbookRole.Accounting, string.Empty, workbook.FullName ?? string.Empty, (worksheet == null) ? string.Empty : (worksheet.CodeName ?? string.Empty));
					_accountingSaveAsService.Execute (context);
				} catch (Exception exception) {
					_logger.Error ("Installment schedule save-as handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingInstallmentSchedule.SaveAsRequested", exception);
				} finally {
					CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (worksheet);
				}
			};
			form.FormClosed += delegate {
				if (_activeInstallmentScheduleForm == form) {
					_activeInstallmentScheduleForm = null;
					_activeInstallmentScheduleWorkbook = null;
				}
				if (_activeInstallmentScheduleOwner == owner) {
					_activeInstallmentScheduleOwner = null;
				}
				if (owner != null) {
					owner.Dispose ();
				}
			};
			form.ShowModeless (owner);
			if (activateExisting) {
				form.FocusInstallmentAmount ();
			}
			_logger.Info ("Accounting installment schedule input form shown. sheet=" + activeSheetCodeName);
		}

		private void HideInstallmentScheduleInput ()
		{
			if (_activeInstallmentScheduleForm != null && !_activeInstallmentScheduleForm.IsDisposed) {
				_activeInstallmentScheduleForm.Hide ();
			}
		}

		private void CloseActiveInstallmentScheduleInput ()
		{
			if (_activeInstallmentScheduleForm == null) {
				return;
			}
			try {
				if (!_activeInstallmentScheduleForm.IsDisposed) {
					_activeInstallmentScheduleForm.Close ();
				}
			} catch {
			} finally {
				_activeInstallmentScheduleForm = null;
				_activeInstallmentScheduleWorkbook = null;
				_activeInstallmentScheduleOwner = null;
			}
		}

		private void ShowPaymentHistoryInput (WorkbookContext context, string activeSheetCodeName)
		{
			EnsurePaymentHistoryInputVisible (context.Workbook, context.Window, activeSheetCodeName, activateExisting: true);
		}

		private void EnsurePaymentHistoryInputVisible (Workbook workbook, Window window, string activeSheetCodeName, bool activateExisting)
		{
			if (!string.Equals (activeSheetCodeName, "お支払い履歴", StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("お支払い履歴シートでのみ利用できます。");
			}
			if (workbook == null || window == null) {
				return;
			}
			AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.LoadFormState (workbook);
			if (_activePaymentHistoryForm != null && !_activePaymentHistoryForm.IsDisposed) {
				ApplySheetAnchoredLocation (_activePaymentHistoryForm, workbook, window, "お支払い履歴", "M1");
				_activePaymentHistoryForm.BindState (state);
				if (!_activePaymentHistoryForm.Visible) {
					_activePaymentHistoryForm.Show ();
				}
				if (activateExisting) {
					_activePaymentHistoryForm.Activate ();
					_activePaymentHistoryForm.FocusReceiptDate ();
				}
				return;
			}
			ExcelWindowOwner owner = ExcelWindowOwner.From (window);
			AccountingPaymentHistoryInputForm form = new AccountingPaymentHistoryInputForm ();
			ApplySheetAnchoredLocation (form, workbook, window, "お支払い履歴", "M1");
			form.BindState (state);
			_activePaymentHistoryForm = form;
			_activePaymentHistoryWorkbook = workbook;
			_activePaymentHistoryOwner = owner;
			form.IssueDateRequested += delegate {
				try {
					AccountingPaymentHistoryFormState state2 = _accountingPaymentHistoryCommandService.ApplyIssueDate (workbook);
					form.BindState (state2);
					form.FocusReceiptDate ();
				} catch (Exception exception) {
					_logger.Error ("Payment history issue date handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.IssueDateRequested", exception);
				}
			};
			form.TodayRequested += delegate {
				try {
					AccountingPaymentHistoryFormState state2 = _accountingPaymentHistoryCommandService.ApplyToday (workbook);
					form.BindState (state2);
					form.FocusReceiptAmount ();
				} catch (Exception exception) {
					_logger.Error ("Payment history today handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.TodayRequested", exception);
				}
			};
			form.AddHistoryRequested += delegate(object sender, AccountingPaymentHistoryEntryRequestEventArgs e) {
				try {
					AccountingPaymentHistoryFormState state2 = _accountingPaymentHistoryCommandService.AddHistoryEntry (workbook, e.Request);
					form.BindState (state2);
					form.FocusReceiptDate ();
				} catch (Exception exception) {
					_logger.Error ("Payment history add handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.AddHistoryRequested", exception);
				}
			};
			form.OutputFutureBalanceRequested += delegate(object sender, AccountingPaymentHistoryEntryRequestEventArgs e) {
				try {
					AccountingPaymentHistoryFormState state2 = _accountingPaymentHistoryCommandService.OutputFutureBalance (workbook, e.Request);
					form.BindState (state2);
					form.FocusReceiptDate ();
				} catch (Exception exception) {
					_logger.Error ("Payment history future balance handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.OutputFutureBalanceRequested", exception);
				}
			};
			form.DeleteBlankRowsRequested += delegate {
				try {
					AccountingPaymentHistoryFormState state2 = _accountingPaymentHistoryCommandService.DeleteBlankReceiptRows (workbook);
					form.BindState (state2);
					form.FocusReceiptDate ();
				} catch (Exception exception) {
					_logger.Error ("Payment history delete blank rows handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.DeleteBlankRowsRequested", exception);
				}
			};
			form.ResetRequested += delegate {
				try {
					DialogResult dialogResult = MessageBox.Show (form, "お支払い履歴は全てクリアされます。よろしいですか？", "案件情報System", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
					if (dialogResult == DialogResult.OK) {
						AccountingPaymentHistoryFormState state2 = _accountingPaymentHistoryCommandService.Reset (workbook);
						form.BindState (state2);
						form.FocusReceiptDate ();
					}
				} catch (Exception exception) {
					_logger.Error ("Payment history reset handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.ResetRequested", exception);
				}
			};
			form.SaveAsRequested += delegate {
				Worksheet worksheet = null;
				try {
					worksheet = workbook.ActiveSheet as Worksheet;
					WorkbookContext context = new WorkbookContext (workbook, window, WorkbookRole.Accounting, string.Empty, workbook.FullName ?? string.Empty, (worksheet == null) ? string.Empty : (worksheet.CodeName ?? string.Empty));
					_accountingSaveAsService.Execute (context);
				} catch (Exception exception) {
					_logger.Error ("Payment history save-as handler failed.", exception);
					_userErrorService.ShowUserError ("AccountingPaymentHistory.SaveAsRequested", exception);
				} finally {
					CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (worksheet);
				}
			};
			form.FormClosed += delegate {
				if (_activePaymentHistoryForm == form) {
					_activePaymentHistoryForm = null;
					_activePaymentHistoryWorkbook = null;
				}
				if (_activePaymentHistoryOwner == owner) {
					_activePaymentHistoryOwner = null;
				}
				if (owner != null) {
					owner.Dispose ();
				}
			};
			form.ShowModeless (owner);
			if (activateExisting) {
				form.FocusReceiptDate ();
			}
			_logger.Info ("Accounting payment history input form shown. sheet=" + activeSheetCodeName);
		}

		private void HidePaymentHistoryInput ()
		{
			if (_activePaymentHistoryForm != null && !_activePaymentHistoryForm.IsDisposed) {
				_activePaymentHistoryForm.Hide ();
			}
		}

		private void CloseActivePaymentHistoryInput ()
		{
			if (_activePaymentHistoryForm == null) {
				return;
			}
			try {
				if (!_activePaymentHistoryForm.IsDisposed) {
					_activePaymentHistoryForm.Close ();
				}
			} catch {
			} finally {
				_activePaymentHistoryForm = null;
				_activePaymentHistoryWorkbook = null;
				_activePaymentHistoryOwner = null;
			}
		}

		private void ApplySheetAnchoredLocation (Form form, Workbook workbook, Window window, string sheetCodeName, string anchorAddress)
		{
			System.Drawing.Point? point = TryCalculateSheetAnchoredLocation (workbook, window, sheetCodeName, anchorAddress);
			if (point.HasValue) {
				form.StartPosition = FormStartPosition.Manual;
				form.Location = point.Value;
			}
		}

		private System.Drawing.Point? TryCalculateSheetAnchoredLocation (Workbook workbook, Window window, string sheetCodeName, string anchorAddress)
		{
			Worksheet worksheet = null;
			Range range = null;
			try {
				if (workbook == null || window == null || string.IsNullOrWhiteSpace (sheetCodeName) || string.IsNullOrWhiteSpace (anchorAddress)) {
					return null;
				}
				worksheet = FindWorksheetByCodeName (workbook, sheetCodeName);
				if (worksheet == null) {
					return null;
				}
				range = ((_Worksheet)worksheet).get_Range ((object)anchorAddress, Type.Missing);
				double num = Convert.ToDouble ((dynamic)window.Zoom);
				double a = Convert.ToDouble ((dynamic)range.Left) * num / 100.0 * 96.0 / 72.0;
				double a2 = Convert.ToDouble ((dynamic)range.Top) * num / 100.0 * 96.0 / 72.0;
				return new System.Drawing.Point (window.PointsToScreenPixelsX (0) + Convert.ToInt32 (Math.Round (a)), window.PointsToScreenPixelsY (0) + Convert.ToInt32 (Math.Round (a2)));
			} catch (Exception exception) {
				_logger.Error ("Accounting form anchored location calculation failed. sheet=" + (sheetCodeName ?? string.Empty), exception);
				return null;
			} finally {
				CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (range);
				CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (worksheet);
			}
		}

		private static Worksheet FindWorksheetByCodeName (Workbook workbook, string sheetCodeName)
		{
			if (workbook == null || string.IsNullOrWhiteSpace (sheetCodeName)) {
				return null;
			}
			Sheets worksheets = null;
			try {
				worksheets = workbook.Worksheets;
				int worksheetCount = (worksheets == null) ? 0 : worksheets.Count;
				for (int index = 1; index <= worksheetCount; index++) {
					Worksheet worksheet = null;
					bool isMatch = false;
					try {
						worksheet = worksheets [index] as Worksheet;
						isMatch = worksheet != null && string.Equals (worksheet.CodeName, sheetCodeName, StringComparison.OrdinalIgnoreCase);
						if (isMatch) {
							return worksheet;
						}
					} finally {
						if (!isMatch) {
							CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (worksheet);
						}
					}
				}
				return null;
			} finally {
				CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (worksheets);
			}
		}

		private void ApplyReverseGoalSeek (Workbook workbook, string activeSheetCodeName, AccountingReverseGoalSeekRequest request)
		{
			if (request == null) {
				throw new InvalidOperationException ("逆算金額を数値で入力してください。");
			}
			Range range = _accountingWorkbookService.TryGetActiveCell (workbook);
			if (!_accountingWorkbookService.IsWithinRange (workbook, activeSheetCodeName, range, "F17:F20")) {
				throw new InvalidOperationException ("逆算対象額を表示したいセルを1つ選択してください。");
			}
			_accountingWorkbookService.ExecuteGoalSeek (workbook, activeSheetCodeName, "F27", range, request.TargetAmount);
			_accountingWorkbookService.RoundDownCell (range, 0);
			double num = _accountingWorkbookService.ReadDouble (range);
			if (num < 0.0) {
				_accountingWorkbookService.WriteCell (workbook, activeSheetCodeName, "B" + range.Row, "実費込");
			}
		}

		private void CloseActiveReverseTool (bool clearHighlight)
		{
			if (_activeReverseToolForm == null) {
				return;
			}
			try {
				if (clearHighlight && _activeReverseToolWorkbook != null && !string.IsNullOrWhiteSpace (_activeReverseToolSheetName)) {
					_accountingWorkbookService.ClearReverseToolTargets (_activeReverseToolWorkbook, _activeReverseToolSheetName);
				}
				if (!_activeReverseToolForm.IsDisposed) {
					_activeReverseToolForm.Close ();
				}
			} catch {
			} finally {
				_activeReverseToolForm = null;
				_activeReverseToolWorkbook = null;
				_activeReverseToolSheetName = string.Empty;
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

		private static string SafeActiveSheetCodeName (Workbook workbook)
		{
			try {
				Worksheet worksheet = ((workbook == null) ? null : (workbook.ActiveSheet as Worksheet));
				return (worksheet == null) ? string.Empty : (worksheet.CodeName ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}

		private static bool IsMainFormSheet (string activeSheetCodeName)
		{
			return string.Equals (activeSheetCodeName, "見積書", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "請求書", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "領収書", StringComparison.OrdinalIgnoreCase) || string.Equals (activeSheetCodeName, "会計依頼書", StringComparison.OrdinalIgnoreCase);
		}

	}
}
