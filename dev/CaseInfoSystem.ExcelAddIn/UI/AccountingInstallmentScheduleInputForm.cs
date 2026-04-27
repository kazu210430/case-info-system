using System;
using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed partial class AccountingInstallmentScheduleInputForm : Form
	{
		private Button btnIssueDate;

		private Button btnReset;

		private Button btnCreateSchedule;

		private Button btnApplyChange;

		private Button btnSaveAs;

		private TextBox txtBillingAmount;

		private TextBox txtExpenseAmount;

		private TextBox txtWithholding;

		private TextBox txtFirstDueDate;

		private TextBox txtDepositAmount;

		private TextBox txtInstallmentAmount;

		private TextBox txtChangeRound;

		private TextBox txtChangedInstallmentAmount;

		internal event EventHandler IssueDateRequested;

		internal event EventHandler ResetRequested;

		internal event EventHandler SaveAsRequested;

		internal event EventHandler<AccountingInstallmentScheduleCreateRequestEventArgs> CreateScheduleRequested;

		internal event EventHandler<AccountingInstallmentScheduleChangeRequestEventArgs> ApplyChangeRequested;

		internal AccountingInstallmentScheduleInputForm ()
		{
			InitializeComponent ();
			AccountingFormButtonAppearanceHelper.Apply (btnIssueDate, btnReset, btnCreateSchedule, btnApplyChange, btnSaveAs);
			ButtonCursorHelper.ApplyHandCursor (this);
		}

		internal void ShowModeless (IWin32Window owner)
		{
			if (owner == null) {
				Show ();
			} else {
				Show (owner);
			}
		}

		internal void BindState (AccountingInstallmentScheduleFormState state)
		{
			txtBillingAmount.Text = ((state == null) ? string.Empty : (state.BillingAmountText ?? string.Empty));
			txtExpenseAmount.Text = ((state == null) ? string.Empty : (state.ExpenseAmountText ?? string.Empty));
			txtWithholding.Text = ((state == null) ? string.Empty : (state.WithholdingText ?? string.Empty));
			txtFirstDueDate.Text = ((state == null) ? string.Empty : (state.FirstDueDateText ?? string.Empty));
			txtDepositAmount.Text = ((state == null) ? string.Empty : (state.DepositAmountText ?? string.Empty));
			txtInstallmentAmount.Text = ((state == null) ? string.Empty : (state.InstallmentAmountText ?? string.Empty));
			txtChangeRound.Text = ((state == null) ? string.Empty : (state.ChangeRoundText ?? string.Empty));
			txtChangedInstallmentAmount.Text = ((state == null) ? string.Empty : (state.ChangedInstallmentAmountText ?? string.Empty));
		}

		internal void FocusInstallmentAmount ()
		{
			txtInstallmentAmount.Focus ();
		}

		private void BtnIssueDate_Click (object sender, EventArgs e)
		{
			this.IssueDateRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnReset_Click (object sender, EventArgs e)
		{
			this.ResetRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnSaveAs_Click (object sender, EventArgs e)
		{
			this.SaveAsRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnCreateSchedule_Click (object sender, EventArgs e)
		{
			this.CreateScheduleRequested?.Invoke (this, new AccountingInstallmentScheduleCreateRequestEventArgs (new AccountingInstallmentScheduleCreateRequest {
				BillingAmountText = txtBillingAmount.Text,
				ExpenseAmountText = txtExpenseAmount.Text,
				WithholdingText = txtWithholding.Text,
				FirstDueDateText = txtFirstDueDate.Text,
				DepositAmountText = txtDepositAmount.Text,
				InstallmentAmountText = txtInstallmentAmount.Text
			}));
		}

		private void ShowPendingMessage (object sender, EventArgs e)
		{
			this.ApplyChangeRequested?.Invoke (this, new AccountingInstallmentScheduleChangeRequestEventArgs (new AccountingInstallmentScheduleChangeRequest {
				BillingAmountText = txtBillingAmount.Text,
				ExpenseAmountText = txtExpenseAmount.Text,
				WithholdingText = txtWithholding.Text,
				ChangeRoundText = txtChangeRound.Text,
				ChangedInstallmentAmountText = txtChangedInstallmentAmount.Text
			}));
		}

		private void ShowInvoiceEditRestrictedMessage (object sender, KeyEventArgs e)
		{
			MessageBox.Show (this, "入力フォームでは変更できません。" + Environment.NewLine + "変更は請求書シートで行ってください", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}

		private static void NumericTextBox_KeyPress (object sender, KeyPressEventArgs e)
		{
			if (!char.IsControl (e.KeyChar) && !char.IsDigit (e.KeyChar)) {
				e.Handled = true;
			}
		}

		private static void CurrencyTextBox_Leave (object sender, EventArgs e)
		{
			if (sender is TextBox textBox) {
				string text = (textBox.Text ?? string.Empty).Replace (",", string.Empty).Trim ();
				long result;
				if (text.Length == 0) {
					textBox.Text = string.Empty;
				} else if (long.TryParse (text, out result)) {
					textBox.Text = result.ToString ("#,##0");
				}
			}
		}

		private void InitializeComponent ()
		{
			base.SuspendLayout ();
			base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
			this.BackColor = System.Drawing.Color.FromArgb (229, 245, 255);
			base.ClientSize = new System.Drawing.Size (540, 505);
			this.Font = new System.Drawing.Font ("Yu Gothic UI", 10f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
			base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.Name = "AccountingInstallmentScheduleInputForm";
			base.ShowInTaskbar = false;
			base.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "分割払い予定表入力フォーム";
			this.btnIssueDate = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateButton ("発行日を入力", new System.Drawing.Point (64, 10), new System.Drawing.Size (120, 33), new System.EventHandler (BtnIssueDate_Click));
			this.btnReset = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateButton ("リセット", new System.Drawing.Point (341, 10), new System.Drawing.Size (120, 33), new System.EventHandler (BtnReset_Click), System.Drawing.Color.FromArgb (255, 192, 192));
			System.Windows.Forms.GroupBox groupBox = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateGroupBox ("請求書の読込内容", new System.Drawing.Point (24, 48), new System.Drawing.Size (475, 150));
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("請求額", new System.Drawing.Point (24, 28), new System.Drawing.Size (72, 20)));
			this.txtBillingAmount = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateReadOnlyTextBox (new System.Drawing.Point (32, 55), new System.Drawing.Size (120, 25));
			this.txtBillingAmount.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			groupBox.Controls.Add (this.txtBillingAmount);
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("円", new System.Drawing.Point (160, 58), new System.Drawing.Size (24, 20)));
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("実費等", new System.Drawing.Point (196, 28), new System.Drawing.Size (72, 20)));
			this.txtExpenseAmount = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateReadOnlyTextBox (new System.Drawing.Point (204, 55), new System.Drawing.Size (120, 25));
			this.txtExpenseAmount.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			groupBox.Controls.Add (this.txtExpenseAmount);
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("円", new System.Drawing.Point (332, 58), new System.Drawing.Size (24, 20)));
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("源泉処理", new System.Drawing.Point (380, 28), new System.Drawing.Size (72, 20)));
			this.txtWithholding = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateReadOnlyTextBox (new System.Drawing.Point (390, 55), new System.Drawing.Size (53, 25));
			this.txtWithholding.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
			this.txtWithholding.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			groupBox.Controls.Add (this.txtWithholding);
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("第1回期限", new System.Drawing.Point (24, 92), new System.Drawing.Size (88, 20)));
			this.txtFirstDueDate = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateReadOnlyTextBox (new System.Drawing.Point (32, 117), new System.Drawing.Size (120, 25));
			this.txtFirstDueDate.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			groupBox.Controls.Add (this.txtFirstDueDate);
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("お預かり金額", new System.Drawing.Point (196, 92), new System.Drawing.Size (120, 20)));
			this.txtDepositAmount = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateReadOnlyTextBox (new System.Drawing.Point (204, 117), new System.Drawing.Size (120, 25));
			this.txtDepositAmount.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			groupBox.Controls.Add (this.txtDepositAmount);
			groupBox.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("円", new System.Drawing.Point (332, 120), new System.Drawing.Size (24, 20)));
			System.Windows.Forms.GroupBox groupBox2 = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateGroupBox ("分割払い作成", new System.Drawing.Point (24, 205), new System.Drawing.Size (475, 92));
			groupBox2.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("分割払い額", new System.Drawing.Point (24, 30), new System.Drawing.Size (90, 20)));
			this.txtInstallmentAmount = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateEditableTextBox (new System.Drawing.Point (32, 56), new System.Drawing.Size (120, 25));
			this.txtInstallmentAmount.KeyPress += new System.Windows.Forms.KeyPressEventHandler (NumericTextBox_KeyPress);
			this.txtInstallmentAmount.Leave += new System.EventHandler (CurrencyTextBox_Leave);
			groupBox2.Controls.Add (this.txtInstallmentAmount);
			groupBox2.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("円", new System.Drawing.Point (160, 59), new System.Drawing.Size (24, 20)));
			this.btnCreateSchedule = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateButton ("この分割払い額で\r\n予定表を作成", new System.Drawing.Point (271, 20), new System.Drawing.Size (184, 56), new System.EventHandler (BtnCreateSchedule_Click), System.Drawing.Color.FromArgb (0, 0, 128), System.Drawing.Color.White);
			groupBox2.Controls.Add (this.btnCreateSchedule);
			System.Windows.Forms.GroupBox groupBox3 = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateGroupBox ("途中の額から変更", new System.Drawing.Point (24, 304), new System.Drawing.Size (475, 144));
			groupBox3.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("途中で変更する回", new System.Drawing.Point (24, 26), new System.Drawing.Size (130, 20)));
			this.txtChangeRound = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateEditableTextBox (new System.Drawing.Point (40, 52), new System.Drawing.Size (37, 25));
			this.txtChangeRound.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtChangeRound.KeyPress += new System.Windows.Forms.KeyPressEventHandler (NumericTextBox_KeyPress);
			groupBox3.Controls.Add (this.txtChangeRound);
			groupBox3.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("回目から変更", new System.Drawing.Point (84, 55), new System.Drawing.Size (96, 20)));
			groupBox3.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("変更後の分割払い額", new System.Drawing.Point (24, 86), new System.Drawing.Size (150, 20)));
			this.txtChangedInstallmentAmount = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateEditableTextBox (new System.Drawing.Point (32, 111), new System.Drawing.Size (120, 25));
			this.txtChangedInstallmentAmount.KeyPress += new System.Windows.Forms.KeyPressEventHandler (NumericTextBox_KeyPress);
			this.txtChangedInstallmentAmount.Leave += new System.EventHandler (CurrencyTextBox_Leave);
			groupBox3.Controls.Add (this.txtChangedInstallmentAmount);
			groupBox3.Controls.Add (CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateLabel ("円", new System.Drawing.Point (160, 114), new System.Drawing.Size (24, 20)));
			this.btnApplyChange = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateButton ("途中の額から変更を\r\n予定表に反映", new System.Drawing.Point (271, 67), new System.Drawing.Size (184, 56), new System.EventHandler (ShowPendingMessage));
			groupBox3.Controls.Add (this.btnApplyChange);
			this.btnSaveAs = CaseInfoSystem.ExcelAddIn.UI.AccountingInstallmentScheduleInputForm.CreateButton ("別名保存", new System.Drawing.Point (64, 457), new System.Drawing.Size (120, 33), new System.EventHandler (BtnSaveAs_Click));
			base.Controls.Add (this.btnIssueDate);
			base.Controls.Add (this.btnReset);
			base.Controls.Add (groupBox);
			base.Controls.Add (groupBox2);
			base.Controls.Add (groupBox3);
			base.Controls.Add (this.btnSaveAs);
			base.ResumeLayout (false);
		}

		private static GroupBox CreateGroupBox (string text, Point location, Size size)
		{
			return new GroupBox {
				Text = text,
				Location = location,
				Size = size
			};
		}

		private static Label CreateLabel (string text, Point location, Size size)
		{
			return new Label {
				AutoSize = false,
				Text = text,
				Location = location,
				Size = size,
				BackColor = Color.Transparent
			};
		}

		private static TextBox CreateReadOnlyTextBox (Point location, Size size)
		{
			return new TextBox {
				Location = location,
				Size = size,
				BorderStyle = BorderStyle.FixedSingle,
				ReadOnly = true,
				BackColor = Color.FromArgb (229, 245, 255),
				TextAlign = HorizontalAlignment.Right
			};
		}

		private static TextBox CreateEditableTextBox (Point location, Size size)
		{
			return new TextBox {
				Location = location,
				Size = size,
				BorderStyle = BorderStyle.FixedSingle,
				TextAlign = HorizontalAlignment.Right
			};
		}

		private static Button CreateButton (string text, Point location, Size size, EventHandler clickHandler)
		{
			return CreateButton (text, location, size, clickHandler, SystemColors.Control, SystemColors.ControlText);
		}

		private static Button CreateButton (string text, Point location, Size size, EventHandler clickHandler, Color backColor)
		{
			return CreateButton (text, location, size, clickHandler, backColor, SystemColors.ControlText);
		}

		private static Button CreateButton (string text, Point location, Size size, EventHandler clickHandler, Color backColor, Color foreColor)
		{
			Button button = new Button {
				Text = text,
				Location = location,
				Size = size,
				BackColor = backColor,
				ForeColor = foreColor,
				UseVisualStyleBackColor = false
			};
			button.Click += clickHandler;
			return button;
		}
	}
}

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingInstallmentScheduleCreateRequestEventArgs : EventArgs
	{
		internal AccountingInstallmentScheduleCreateRequest Request { get; }

		internal AccountingInstallmentScheduleCreateRequestEventArgs (AccountingInstallmentScheduleCreateRequest request)
		{
			Request = request ?? new AccountingInstallmentScheduleCreateRequest ();
		}
	}
}

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingInstallmentScheduleChangeRequestEventArgs : EventArgs
	{
		internal AccountingInstallmentScheduleChangeRequest Request { get; }

		internal AccountingInstallmentScheduleChangeRequestEventArgs (AccountingInstallmentScheduleChangeRequest request)
		{
			Request = request ?? new AccountingInstallmentScheduleChangeRequest ();
		}
	}
}
