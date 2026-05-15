using System;
using System.Drawing;
using System.Globalization;
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

		private const string ProductTitle = "案件情報System";

		private const string CloseThroughFormMessage = "フォームの「Excelを閉じる」ボタンから閉じてください。";

		private const string InstallmentScheduleFormKind = "InstallmentSchedule";

		private const string PaymentHistoryFormKind = "PaymentHistory";

		private const string ReverseGoalSeekFormKind = "ReverseGoalSeek";

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly AccountingInstallmentScheduleCommandService _accountingInstallmentScheduleCommandService;

		private readonly AccountingPaymentHistoryCommandService _accountingPaymentHistoryCommandService;

		private readonly AccountingSaveAsService _accountingSaveAsService;

		private readonly UserErrorService _userErrorService;

		private readonly Logger _logger;

		private AccountingReverseGoalSeekForm _activeReverseToolForm;

		private Workbook _activeReverseToolWorkbook;

		private string _activeReverseToolSheetName;

		private ExcelWindowOwner _activeReverseToolOwner;

		private bool _activeReverseToolHighlightCleanupCompleted;

		private AccountingInstallmentScheduleInputForm _activeInstallmentScheduleForm;

		private Workbook _activeInstallmentScheduleWorkbook;

		private ExcelWindowOwner _activeInstallmentScheduleOwner;

		private AccountingPaymentHistoryInputForm _activePaymentHistoryForm;

		private Workbook _activePaymentHistoryWorkbook;

		private ExcelWindowOwner _activePaymentHistoryOwner;

		private string _formButtonWorkbookCloseKey;

		private string _formButtonWorkbookCloseFormKind;

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
			case AccountingNavigationActionIds.SetInstallmentScheduleIssueDate:
				EnsureInstallmentScheduleActionSheet (text);
				ApplyInstallmentScheduleIssueDate (context.Workbook, GetActiveInstallmentScheduleFormForWorkbook (context.Workbook));
				break;
			case AccountingNavigationActionIds.ResetInstallmentSchedule:
				EnsureInstallmentScheduleActionSheet (text);
				ResetInstallmentSchedule (context.Workbook, GetActiveInstallmentScheduleFormForWorkbook (context.Workbook));
				break;
			case AccountingNavigationActionIds.SetPaymentHistoryIssueDate:
				EnsurePaymentHistoryActionSheet (text);
				ApplyPaymentHistoryIssueDate (context.Workbook, GetActivePaymentHistoryFormForWorkbook (context.Workbook));
				break;
			case AccountingNavigationActionIds.DeleteSelectedPaymentHistoryRows:
				EnsurePaymentHistoryActionSheet (text);
				DeleteSelectedPaymentHistoryRows (context.Workbook, GetActivePaymentHistoryFormForWorkbook (context.Workbook));
				break;
			case AccountingNavigationActionIds.ResetPaymentHistory:
				EnsurePaymentHistoryActionSheet (text);
				ResetPaymentHistory (context.Workbook, GetActivePaymentHistoryFormForWorkbook (context.Workbook));
				break;
			}
		}

		internal void HandleSheetActivated (Workbook workbook, Window window, string activeSheetCodeName)
		{
			if (workbook == null || _activeReverseToolWorkbook != workbook || !string.Equals (_activeReverseToolSheetName, activeSheetCodeName, StringComparison.OrdinalIgnoreCase) || !IsMainFormSheet (activeSheetCodeName)) {
				CloseActiveReverseTool ();
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
					CloseActiveReverseTool ();
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
			double invoicePaymentAmount = ReadRequiredDouble (workbook, activeSheetCodeName, "F31", "お支払い頂く金額", "AccountingFormHelper.ApplyIssueDateAndDueDate");
			_accountingWorkbookService.WriteCellValue (workbook, activeSheetCodeName, "A1", dateTime);
			if (AccountingIssueDateDueDatePolicy.ShouldWriteNoPaymentNotice (invoicePaymentAmount)) {
				_accountingWorkbookService.ClearValidation (workbook, activeSheetCodeName, "G10");
				_accountingWorkbookService.WriteCellValue (workbook, activeSheetCodeName, "G10", AccountingIssueDateDueDatePolicy.NoPaymentNoticeText);
				_logger.Info ("Accounting issue date applied with no payment notice. sheet=" + activeSheetCodeName + ", issueDate=" + dateTime.ToString ("yyyy-MM-dd"));
				return;
			}
			_accountingWorkbookService.WriteCellValue (workbook, activeSheetCodeName, "G10", dateTime.AddDays (14.0));
			_logger.Info ("Accounting issue date and due date applied. sheet=" + activeSheetCodeName + ", issueDate=" + dateTime.ToString ("yyyy-MM-dd"));
		}

		private double ReadRequiredDouble (Workbook workbook, string sheetName, string address, string itemName, string procedureName)
		{
			object cellValue = _accountingWorkbookService.ReadCellValue (workbook, sheetName, address);
			string displayText = _accountingWorkbookService.ReadDisplayText (workbook, sheetName, address);
			if (AccountingNumericCellReader.TryParseNumericCell (cellValue, displayText, out var value, out var isBlank)) {
				return value;
			}
			InvalidOperationException ex = AccountingNumericCellReader.CreateReadFailureException (sheetName, address, itemName, procedureName, displayText, allowBlankAsZero: false);
			string text = Convert.ToString (cellValue, CultureInfo.InvariantCulture) ?? string.Empty;
			_logger.Error ("Accounting numeric cell read failed. sheet=" + sheetName + ", address=" + address + ", item=" + itemName + ", procedure=" + procedureName + ", displayText=" + (string.IsNullOrWhiteSpace (displayText) ? "（空欄）" : displayText.Trim ()) + ", cellValue=" + text, ex);
			throw ex;
		}

		private void ShowReverseTool (WorkbookContext context, string activeSheetCodeName)
		{
			CloseActiveReverseTool ();
			_activeReverseToolHighlightCleanupCompleted = false;
			_accountingWorkbookService.HighlightReverseToolTargets (context.Workbook, activeSheetCodeName);
			ExcelWindowOwner owner = ExcelWindowOwner.From (context.Window);
			AccountingReverseGoalSeekForm form = new AccountingReverseGoalSeekForm ();
			ApplySheetAnchoredLocation (form, context.Workbook, context.Window, activeSheetCodeName, "AA1");
			_activeReverseToolForm = form;
			_activeReverseToolWorkbook = context.Workbook;
			_activeReverseToolSheetName = activeSheetCodeName;
			_activeReverseToolOwner = owner;
			AttachReverseToolHandlers (form);
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
			AttachInstallmentScheduleInputHandlers (form);
			form.ShowModeless (owner);
			if (activateExisting) {
				form.FocusInstallmentAmount ();
			}
			_logger.Info ("Accounting installment schedule input form shown. sheet=" + activeSheetCodeName);
		}

		private static void EnsureInstallmentScheduleActionSheet (string activeSheetCodeName)
		{
			if (!string.Equals (activeSheetCodeName, AccountingSetSpec.InstallmentSheetName, StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("分割払い予定表シートでのみ利用できます。");
			}
		}

		private void ApplyInstallmentScheduleIssueDate (Workbook workbook, AccountingInstallmentScheduleInputForm form)
		{
			AccountingInstallmentScheduleFormState state = _accountingInstallmentScheduleCommandService.ApplyIssueDate (workbook);
			BindInstallmentScheduleFormState (form, state, focusInstallmentAmount: false);
		}

		private void ResetInstallmentSchedule (Workbook workbook, AccountingInstallmentScheduleInputForm form)
		{
			AccountingInstallmentScheduleFormState state = _accountingInstallmentScheduleCommandService.Reset (workbook);
			BindInstallmentScheduleFormState (form, state, focusInstallmentAmount: true);
		}

		private AccountingInstallmentScheduleInputForm GetActiveInstallmentScheduleFormForWorkbook (Workbook workbook)
		{
			if (_activeInstallmentScheduleForm == null || _activeInstallmentScheduleForm.IsDisposed || !IsSameWorkbook (_activeInstallmentScheduleWorkbook, workbook)) {
				return null;
			}
			return _activeInstallmentScheduleForm;
		}

		private static void BindInstallmentScheduleFormState (AccountingInstallmentScheduleInputForm form, AccountingInstallmentScheduleFormState state, bool focusInstallmentAmount)
		{
			if (form == null || form.IsDisposed) {
				return;
			}
			form.BindState (state);
			if (focusInstallmentAmount) {
				form.FocusInstallmentAmount ();
			}
		}

		private void HideInstallmentScheduleInput ()
		{
			CloseActiveInstallmentScheduleInput ();
		}

		private void CloseActiveInstallmentScheduleInput ()
		{
			AccountingInstallmentScheduleInputForm form = _activeInstallmentScheduleForm;
			if (form == null) {
				return;
			}
			try {
				_logger.Info ("Accounting installment schedule input form close/dispose starting.");
				DetachInstallmentScheduleInputHandlers (form);
				if (!form.IsDisposed) {
					form.Close ();
					if (!form.IsDisposed) {
						form.Dispose ();
					}
				}
				_logger.Info ("Accounting installment schedule input form close/dispose completed.");
			} catch {
			} finally {
				ClearActiveInstallmentScheduleInputReferences ();
				_logger.Info ("Accounting installment schedule input form active references cleared.");
			}
		}

		internal bool TryCancelWorkbookCloseForActiveAccountingForm (Workbook workbook)
		{
			string formKind = GetActiveAccountingFormKindForWorkbook (workbook);
			if (string.IsNullOrWhiteSpace (formKind)) {
				return false;
			}
			if (IsFormButtonWorkbookCloseAllowed (workbook, formKind)) {
				_logger.Info ("Accounting workbook close guard bypassed for form button. formKind=" + formKind + ", cancel=False");
				return false;
			}
			_logger.Info ("Accounting workbook close canceled because accounting form is active. formKind=" + formKind + ", cancel=True");
			ShowCloseThroughFormMessage (formKind);
			return true;
		}

		private string GetActiveAccountingFormKindForWorkbook (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			if (_activeInstallmentScheduleForm != null && !_activeInstallmentScheduleForm.IsDisposed && IsSameWorkbook (_activeInstallmentScheduleWorkbook, workbook)) {
				return InstallmentScheduleFormKind;
			}
			if (_activePaymentHistoryForm != null && !_activePaymentHistoryForm.IsDisposed && IsSameWorkbook (_activePaymentHistoryWorkbook, workbook)) {
				return PaymentHistoryFormKind;
			}
			if (_activeReverseToolForm != null && !_activeReverseToolForm.IsDisposed && IsSameWorkbook (_activeReverseToolWorkbook, workbook)) {
				return ReverseGoalSeekFormKind;
			}
			return string.Empty;
		}

		private bool IsFormButtonWorkbookCloseAllowed (Workbook workbook, string formKind)
		{
			if (string.IsNullOrWhiteSpace (_formButtonWorkbookCloseKey) || string.IsNullOrWhiteSpace (formKind)) {
				return false;
			}
			return string.Equals (_formButtonWorkbookCloseFormKind ?? string.Empty, formKind, StringComparison.OrdinalIgnoreCase)
				&& string.Equals (_formButtonWorkbookCloseKey, SafeWorkbookKey (workbook), StringComparison.OrdinalIgnoreCase);
		}

		private void ShowCloseThroughFormMessage (string formKind)
		{
			IWin32Window owner = ResolveActiveAccountingFormWindow (formKind);
			if (owner == null) {
				MessageBox.Show (CloseThroughFormMessage, ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
			} else {
				MessageBox.Show (owner, CloseThroughFormMessage, ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			_logger.Info ("Accounting workbook close guard message shown. formKind=" + (formKind ?? string.Empty));
		}

		private IWin32Window ResolveActiveAccountingFormWindow (string formKind)
		{
			if (string.Equals (formKind, InstallmentScheduleFormKind, StringComparison.OrdinalIgnoreCase)
				&& _activeInstallmentScheduleForm != null
				&& !_activeInstallmentScheduleForm.IsDisposed) {
				return _activeInstallmentScheduleForm;
			}
			if (string.Equals (formKind, PaymentHistoryFormKind, StringComparison.OrdinalIgnoreCase)
				&& _activePaymentHistoryForm != null
				&& !_activePaymentHistoryForm.IsDisposed) {
				return _activePaymentHistoryForm;
			}
			if (string.Equals (formKind, ReverseGoalSeekFormKind, StringComparison.OrdinalIgnoreCase)
				&& _activeReverseToolForm != null
				&& !_activeReverseToolForm.IsDisposed) {
				return _activeReverseToolForm;
			}
			return null;
		}

		private void RequestWorkbookCloseFromAccountingForm (Workbook workbook, string formKind)
		{
			if (workbook == null) {
				_logger.Info ("Accounting form Excel close request ignored. reason=WorkbookMissing, formKind=" + (formKind ?? string.Empty));
				return;
			}
			string workbookKey = SafeWorkbookKey (workbook);
			Microsoft.Office.Interop.Excel.Application application = TryGetWorkbookApplication (workbook);
			_logger.Info ("Accounting form Excel close button clicked. formKind=" + (formKind ?? string.Empty) + ", workbook=" + workbookKey);
			_formButtonWorkbookCloseKey = workbookKey;
			_formButtonWorkbookCloseFormKind = formKind ?? string.Empty;
			_logger.Info ("Accounting form workbook close allow flag set. formKind=" + (formKind ?? string.Empty) + ", workbook=" + workbookKey);
			try {
				if (string.Equals (formKind, InstallmentScheduleFormKind, StringComparison.OrdinalIgnoreCase)) {
					CloseActiveInstallmentScheduleInput ();
				} else if (string.Equals (formKind, PaymentHistoryFormKind, StringComparison.OrdinalIgnoreCase)) {
					CloseActivePaymentHistoryInput ();
				} else if (string.Equals (formKind, ReverseGoalSeekFormKind, StringComparison.OrdinalIgnoreCase)) {
					CloseActiveReverseTool ();
				}
				_logger.Info ("Accounting form closed before workbook close. formKind=" + (formKind ?? string.Empty) + ", workbook=" + workbookKey);
				_logger.Info ("Accounting form invoking workbook.Close. formKind=" + (formKind ?? string.Empty) + ", workbook=" + workbookKey + ", cancelTouched=False, saveInvoked=False, saveAsInvoked=False, savedForced=False");
				workbook.Close ();
				QuitExcelIfNoWorkbooksAfterFormButtonClose (application, workbookKey, formKind);
			} catch (Exception exception) {
				_logger.Error ("Accounting form workbook close request failed. formKind=" + (formKind ?? string.Empty) + ", workbook=" + workbookKey, exception);
				_userErrorService.ShowUserError ("AccountingForm.ExcelCloseRequested", exception);
			} finally {
				_formButtonWorkbookCloseKey = string.Empty;
				_formButtonWorkbookCloseFormKind = string.Empty;
				_logger.Info ("Accounting form workbook close allow flag cleared. formKind=" + (formKind ?? string.Empty) + ", workbook=" + workbookKey);
			}
		}

		private void QuitExcelIfNoWorkbooksAfterFormButtonClose (Microsoft.Office.Interop.Excel.Application application, string workbookKey, string formKind)
		{
			if (application == null) {
				_logger.Info ("Accounting form button close quit skipped. reason=ApplicationMissing, formKind=" + (formKind ?? string.Empty) + ", workbook=" + (workbookKey ?? string.Empty));
				return;
			}
			bool readFailed;
			int workbooksCount = ReadWorkbooksCount (application, out readFailed);
			string visible = SafeApplicationVisible (application);
			if (readFailed) {
				_logger.Info ("Accounting form button close quit skipped. reason=WorkbooksCountReadFailed, formKind=" + (formKind ?? string.Empty) + ", workbook=" + (workbookKey ?? string.Empty) + ", applicationVisible=" + visible);
				return;
			}
			if (workbooksCount != 0) {
				_logger.Info ("Accounting form button close quit skipped. reason=WorkbookStillOpenOrOtherWorkbookPresent, formKind=" + (formKind ?? string.Empty) + ", workbook=" + (workbookKey ?? string.Empty) + ", workbooksCount=" + workbooksCount.ToString (CultureInfo.InvariantCulture) + ", applicationVisible=" + visible);
				return;
			}
			_logger.Info ("form-button-close-after-workbook-close-no-workbooks-quit. formKind=" + (formKind ?? string.Empty) + ", workbook=" + (workbookKey ?? string.Empty) + ", workbooksCount=0, applicationVisible=" + visible + ", applicationVisibleFalseTouched=False, saveInvoked=False, saveAsInvoked=False, savedForced=False");
			application.Quit ();
		}

		private static Microsoft.Office.Interop.Excel.Application TryGetWorkbookApplication (Workbook workbook)
		{
			try {
				return workbook == null ? null : workbook.Application;
			} catch {
				return null;
			}
		}

		private static int ReadWorkbooksCount (Microsoft.Office.Interop.Excel.Application application, out bool readFailed)
		{
			readFailed = false;
			try {
				return application == null || application.Workbooks == null ? -1 : application.Workbooks.Count;
			} catch {
				readFailed = true;
				return -1;
			}
		}

		private static string SafeApplicationVisible (Microsoft.Office.Interop.Excel.Application application)
		{
			try {
				return application == null ? string.Empty : application.Visible.ToString ();
			} catch {
				return string.Empty;
			}
		}

		private void AttachInstallmentScheduleInputHandlers (AccountingInstallmentScheduleInputForm form)
		{
			if (form == null) {
				return;
			}
			form.CreateScheduleRequested += ActiveInstallmentScheduleForm_CreateScheduleRequested;
			form.ApplyChangeRequested += ActiveInstallmentScheduleForm_ApplyChangeRequested;
			form.ExcelCloseRequested += ActiveInstallmentScheduleForm_ExcelCloseRequested;
			form.FormClosed += ActiveInstallmentScheduleForm_FormClosed;
		}

		private void DetachInstallmentScheduleInputHandlers (AccountingInstallmentScheduleInputForm form)
		{
			if (form == null) {
				return;
			}
			form.CreateScheduleRequested -= ActiveInstallmentScheduleForm_CreateScheduleRequested;
			form.ApplyChangeRequested -= ActiveInstallmentScheduleForm_ApplyChangeRequested;
			form.ExcelCloseRequested -= ActiveInstallmentScheduleForm_ExcelCloseRequested;
			form.FormClosed -= ActiveInstallmentScheduleForm_FormClosed;
			form.ClearRequestHandlers ();
		}

		private void ActiveInstallmentScheduleForm_ExcelCloseRequested (object sender, EventArgs e)
		{
			RequestWorkbookCloseFromAccountingForm (_activeInstallmentScheduleWorkbook, InstallmentScheduleFormKind);
		}

		private void ActiveInstallmentScheduleForm_CreateScheduleRequested (object sender, AccountingInstallmentScheduleCreateRequestEventArgs e)
		{
			if (_activeInstallmentScheduleWorkbook == null || _activeInstallmentScheduleForm == null || _activeInstallmentScheduleForm.IsDisposed) {
				return;
			}
			try {
				AccountingInstallmentScheduleFormState state = _accountingInstallmentScheduleCommandService.CreateSchedule (_activeInstallmentScheduleWorkbook, e.Request);
				BindInstallmentScheduleFormState (_activeInstallmentScheduleForm, state, focusInstallmentAmount: true);
			} catch (Exception exception) {
				_logger.Error ("Installment schedule create handler failed.", exception);
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.CreateScheduleRequested", exception);
			}
		}

		private void ActiveInstallmentScheduleForm_ApplyChangeRequested (object sender, AccountingInstallmentScheduleChangeRequestEventArgs e)
		{
			if (_activeInstallmentScheduleWorkbook == null || _activeInstallmentScheduleForm == null || _activeInstallmentScheduleForm.IsDisposed) {
				return;
			}
			try {
				AccountingInstallmentScheduleFormState state = _accountingInstallmentScheduleCommandService.ApplyChange (_activeInstallmentScheduleWorkbook, e.Request);
				BindInstallmentScheduleFormState (_activeInstallmentScheduleForm, state, focusInstallmentAmount: true);
			} catch (Exception exception) {
				_logger.Error ("Installment schedule apply change handler failed.", exception);
				_userErrorService.ShowUserError ("AccountingInstallmentSchedule.ApplyChangeRequested", exception);
			}
		}

		private void ActiveInstallmentScheduleForm_FormClosed (object sender, FormClosedEventArgs e)
		{
			AccountingInstallmentScheduleInputForm form = sender as AccountingInstallmentScheduleInputForm;
			DetachInstallmentScheduleInputHandlers (form);
			if (ReferenceEquals (form, _activeInstallmentScheduleForm)) {
				ClearActiveInstallmentScheduleInputReferences ();
			}
		}

		private void ClearActiveInstallmentScheduleInputReferences ()
		{
			_activeInstallmentScheduleForm = null;
			_activeInstallmentScheduleWorkbook = null;
			ExcelWindowOwner owner = _activeInstallmentScheduleOwner;
			_activeInstallmentScheduleOwner = null;
			if (owner != null) {
				owner.Dispose ();
			}
		}

		private void ShowPaymentHistoryInput (WorkbookContext context, string activeSheetCodeName)
		{
			EnsurePaymentHistoryInputVisible (context.Workbook, context.Window, activeSheetCodeName, activateExisting: true);
		}

		private static void EnsurePaymentHistoryActionSheet (string activeSheetCodeName)
		{
			if (!string.Equals (activeSheetCodeName, AccountingSetSpec.PaymentHistorySheetName, StringComparison.OrdinalIgnoreCase)) {
				throw new InvalidOperationException ("お支払い履歴シートでのみ利用できます。");
			}
		}

		private void ApplyPaymentHistoryIssueDate (Workbook workbook, AccountingPaymentHistoryInputForm form)
		{
			AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.ApplyIssueDate (workbook);
			BindPaymentHistoryFormState (form, state, focusReceiptDate: true);
		}

		private void ResetPaymentHistory (Workbook workbook, AccountingPaymentHistoryInputForm form)
		{
			DialogResult dialogResult = ShowPaymentHistoryResetConfirmation (form);
			if (dialogResult != DialogResult.OK) {
				return;
			}
			AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.Reset (workbook);
			BindPaymentHistoryFormState (form, state, focusReceiptDate: true);
		}

		private void DeleteSelectedPaymentHistoryRows (Workbook workbook, AccountingPaymentHistoryInputForm form)
		{
			AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.DeleteSelectedRows (workbook);
			BindPaymentHistoryFormState (form, state, focusReceiptDate: true);
		}

		private AccountingPaymentHistoryInputForm GetActivePaymentHistoryFormForWorkbook (Workbook workbook)
		{
			if (_activePaymentHistoryForm == null || _activePaymentHistoryForm.IsDisposed || !IsSameWorkbook (_activePaymentHistoryWorkbook, workbook)) {
				return null;
			}
			return _activePaymentHistoryForm;
		}

		private static void BindPaymentHistoryFormState (AccountingPaymentHistoryInputForm form, AccountingPaymentHistoryFormState state, bool focusReceiptDate)
		{
			if (form == null || form.IsDisposed) {
				return;
			}
			form.BindState (state);
			if (focusReceiptDate) {
				form.FocusReceiptDate ();
			}
		}

		private static DialogResult ShowPaymentHistoryResetConfirmation (IWin32Window owner)
		{
			const string message = "お支払い履歴は全てクリアされます。よろしいですか？";
			const string caption = "案件情報System";
			if (owner == null) {
				return MessageBox.Show (message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
			}
			return MessageBox.Show (owner, message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
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
			AttachPaymentHistoryInputHandlers (form);
			form.ShowModeless (owner);
			if (activateExisting) {
				form.FocusReceiptDate ();
			}
			_logger.Info ("Accounting payment history input form shown. sheet=" + activeSheetCodeName);
		}

		private void HidePaymentHistoryInput ()
		{
			CloseActivePaymentHistoryInput ();
		}

		private void CloseActivePaymentHistoryInput ()
		{
			AccountingPaymentHistoryInputForm form = _activePaymentHistoryForm;
			if (form == null) {
				return;
			}
			try {
				_logger.Info ("Accounting payment history input form close/dispose starting.");
				DetachPaymentHistoryInputHandlers (form);
				if (!form.IsDisposed) {
					form.Close ();
					if (!form.IsDisposed) {
						form.Dispose ();
					}
				}
				_logger.Info ("Accounting payment history input form close/dispose completed.");
			} catch {
			} finally {
				ClearActivePaymentHistoryInputReferences ();
				_logger.Info ("Accounting payment history input form active references cleared.");
			}
		}

		private void AttachPaymentHistoryInputHandlers (AccountingPaymentHistoryInputForm form)
		{
			if (form == null) {
				return;
			}
			form.TodayRequested += ActivePaymentHistoryForm_TodayRequested;
			form.AddHistoryRequested += ActivePaymentHistoryForm_AddHistoryRequested;
			form.OutputFutureBalanceRequested += ActivePaymentHistoryForm_OutputFutureBalanceRequested;
			form.ExcelCloseRequested += ActivePaymentHistoryForm_ExcelCloseRequested;
			form.FormClosed += ActivePaymentHistoryForm_FormClosed;
		}

		private void DetachPaymentHistoryInputHandlers (AccountingPaymentHistoryInputForm form)
		{
			if (form == null) {
				return;
			}
			form.TodayRequested -= ActivePaymentHistoryForm_TodayRequested;
			form.AddHistoryRequested -= ActivePaymentHistoryForm_AddHistoryRequested;
			form.OutputFutureBalanceRequested -= ActivePaymentHistoryForm_OutputFutureBalanceRequested;
			form.ExcelCloseRequested -= ActivePaymentHistoryForm_ExcelCloseRequested;
			form.FormClosed -= ActivePaymentHistoryForm_FormClosed;
			form.ClearRequestHandlers ();
		}

		private void ActivePaymentHistoryForm_ExcelCloseRequested (object sender, EventArgs e)
		{
			RequestWorkbookCloseFromAccountingForm (_activePaymentHistoryWorkbook, PaymentHistoryFormKind);
		}

		private void ActivePaymentHistoryForm_TodayRequested (object sender, EventArgs e)
		{
			if (_activePaymentHistoryWorkbook == null || _activePaymentHistoryForm == null || _activePaymentHistoryForm.IsDisposed) {
				return;
			}
			try {
				AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.ApplyToday (_activePaymentHistoryWorkbook);
				BindPaymentHistoryFormState (_activePaymentHistoryForm, state, focusReceiptDate: false);
				_activePaymentHistoryForm.FocusReceiptAmount ();
			} catch (Exception exception) {
				_logger.Error ("Payment history today handler failed.", exception);
				_userErrorService.ShowUserError ("AccountingPaymentHistory.TodayRequested", exception);
			}
		}

		private void ActivePaymentHistoryForm_AddHistoryRequested (object sender, AccountingPaymentHistoryEntryRequestEventArgs e)
		{
			if (_activePaymentHistoryWorkbook == null || _activePaymentHistoryForm == null || _activePaymentHistoryForm.IsDisposed) {
				return;
			}
			try {
				AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.AddHistoryEntry (_activePaymentHistoryWorkbook, e.Request);
				BindPaymentHistoryFormState (_activePaymentHistoryForm, state, focusReceiptDate: true);
			} catch (Exception exception) {
				_logger.Error ("Payment history add handler failed.", exception);
				_userErrorService.ShowUserError ("AccountingPaymentHistory.AddHistoryRequested", exception);
			}
		}

		private void ActivePaymentHistoryForm_OutputFutureBalanceRequested (object sender, AccountingPaymentHistoryEntryRequestEventArgs e)
		{
			if (_activePaymentHistoryWorkbook == null || _activePaymentHistoryForm == null || _activePaymentHistoryForm.IsDisposed) {
				return;
			}
			try {
				AccountingPaymentHistoryFormState state = _accountingPaymentHistoryCommandService.OutputFutureBalance (_activePaymentHistoryWorkbook, e.Request);
				BindPaymentHistoryFormState (_activePaymentHistoryForm, state, focusReceiptDate: true);
			} catch (Exception exception) {
				_logger.Error ("Payment history future balance handler failed.", exception);
				_userErrorService.ShowUserError ("AccountingPaymentHistory.OutputFutureBalanceRequested", exception);
			}
		}

		private void ActivePaymentHistoryForm_FormClosed (object sender, FormClosedEventArgs e)
		{
			AccountingPaymentHistoryInputForm form = sender as AccountingPaymentHistoryInputForm;
			DetachPaymentHistoryInputHandlers (form);
			if (ReferenceEquals (form, _activePaymentHistoryForm)) {
				ClearActivePaymentHistoryInputReferences ();
			}
		}

		private void ClearActivePaymentHistoryInputReferences ()
		{
			_activePaymentHistoryForm = null;
			_activePaymentHistoryWorkbook = null;
			ExcelWindowOwner owner = _activePaymentHistoryOwner;
			_activePaymentHistoryOwner = null;
			if (owner != null) {
				owner.Dispose ();
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

		private void AttachReverseToolHandlers (AccountingReverseGoalSeekForm form)
		{
			if (form == null) {
				return;
			}
			form.Confirmed += ActiveReverseToolForm_Confirmed;
			form.ExcelCloseRequested += ActiveReverseToolForm_ExcelCloseRequested;
			form.FormClosing += ActiveReverseToolForm_FormClosing;
			form.FormClosed += ActiveReverseToolForm_FormClosed;
		}

		private void DetachReverseToolHandlers (AccountingReverseGoalSeekForm form)
		{
			if (form == null) {
				return;
			}
			form.Confirmed -= ActiveReverseToolForm_Confirmed;
			form.ExcelCloseRequested -= ActiveReverseToolForm_ExcelCloseRequested;
			form.FormClosing -= ActiveReverseToolForm_FormClosing;
			form.FormClosed -= ActiveReverseToolForm_FormClosed;
			form.ClearRequestHandlers ();
		}

		private void ActiveReverseToolForm_ExcelCloseRequested (object sender, EventArgs e)
		{
			RequestWorkbookCloseFromAccountingForm (_activeReverseToolWorkbook, ReverseGoalSeekFormKind);
		}

		private void ActiveReverseToolForm_Confirmed (object sender, AccountingReverseGoalSeekConfirmedEventArgs e)
		{
			if (_activeReverseToolWorkbook == null || string.IsNullOrWhiteSpace (_activeReverseToolSheetName)) {
				return;
			}
			try {
				ApplyReverseGoalSeek (_activeReverseToolWorkbook, _activeReverseToolSheetName, e.Request);
			} catch (Exception exception) {
				_logger.Error ("Accounting reverse tool confirmed handler failed.", exception);
				_userErrorService.ShowUserError ("AccountingReverseGoalSeek", exception);
			}
		}

		private void ActiveReverseToolForm_FormClosing (object sender, FormClosingEventArgs e)
		{
			if (ReferenceEquals (sender, _activeReverseToolForm)) {
				CleanupActiveReverseToolHighlightOnce ("FormClosing");
			}
		}

		private void ActiveReverseToolForm_FormClosed (object sender, FormClosedEventArgs e)
		{
			AccountingReverseGoalSeekForm form = sender as AccountingReverseGoalSeekForm;
			if (ReferenceEquals (form, _activeReverseToolForm)) {
				CleanupActiveReverseToolHighlightOnce ("FormClosed");
			}
			DetachReverseToolHandlers (form);
			if (ReferenceEquals (form, _activeReverseToolForm)) {
				ClearActiveReverseToolReferences ();
				_logger.Info ("Accounting reverse tool active references cleared.");
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

		private void CloseActiveReverseTool ()
		{
			AccountingReverseGoalSeekForm form = _activeReverseToolForm;
			if (form == null) {
				return;
			}
			try {
				_logger.Info ("Accounting reverse tool form close/dispose starting.");
				if (!form.IsDisposed) {
					form.Close ();
					if (!form.IsDisposed) {
						form.Dispose ();
					}
				}
				_logger.Info ("Accounting reverse tool form close/dispose completed.");
			} catch {
			} finally {
				CleanupActiveReverseToolHighlightOnce ("CloseActiveReverseTool");
				if (ReferenceEquals (form, _activeReverseToolForm)) {
					DetachReverseToolHandlers (form);
					ClearActiveReverseToolReferences ();
					_logger.Info ("Accounting reverse tool active references cleared.");
				}
			}
		}

		private void CleanupActiveReverseToolHighlightOnce (string reason)
		{
			if (_activeReverseToolHighlightCleanupCompleted) {
				return;
			}
			_activeReverseToolHighlightCleanupCompleted = true;
			ClearReverseToolHighlight ();
			_logger.Info ("Accounting reverse tool highlight cleanup completed. reason=" + (reason ?? string.Empty));
		}

		private void ClearReverseToolHighlight ()
		{
			try {
				if (_activeReverseToolWorkbook != null && !string.IsNullOrWhiteSpace (_activeReverseToolSheetName)) {
					_accountingWorkbookService.ClearReverseToolTargets (_activeReverseToolWorkbook, _activeReverseToolSheetName);
				}
			} catch (Exception exception) {
				_logger.Error ("Accounting reverse tool highlight cleanup failed.", exception);
			}
		}

		private void ClearActiveReverseToolReferences ()
		{
			_activeReverseToolForm = null;
			_activeReverseToolWorkbook = null;
			_activeReverseToolSheetName = string.Empty;
			ExcelWindowOwner owner = _activeReverseToolOwner;
			_activeReverseToolOwner = null;
			if (owner != null) {
				owner.Dispose ();
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

		private static string SafeWorkbookKey (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			try {
				string fullName = workbook.FullName ?? string.Empty;
				return string.IsNullOrWhiteSpace (fullName) ? (workbook.Name ?? string.Empty) : fullName;
			} catch {
				return string.Empty;
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
