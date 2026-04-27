using System;
using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed partial class AccountingPaymentHistoryInputForm : Form
	{
		private Button btn発行日を入力;

		private Button btnリセット;

		private VbaFramePanel frame請求書の記載内容;

		private Label lbl請求額;

		private TextBox text請求額;

		private Label lbl請求額円;

		private Label lbl実費等;

		private TextBox text実費等;

		private Label lbl実費等円;

		private Label lbl源泉処理;

		private TextBox text源泉処理;

		private Label lblお預かり金額;

		private TextBox textお預かり金額;

		private Label lblお預かり金額円;

		private VbaFramePanel frame領収内容;

		private Label lbl領収日;

		private TextBox text領収日;

		private Button btn今日;

		private Label lbl領収額;

		private TextBox text領収額;

		private Label lbl領収額円;

		private Button btn履歴を入力;

		private Button btn今後の残高推移を出力;

		private Button btn別名保存;

		private Label lbl修正削除見出し;

		private Label lbl手順1;

		private Label lbl手順2;

		private Button btn削除;

		private Label lbl手順3;

		internal event EventHandler IssueDateRequested;

		internal event EventHandler TodayRequested;

		internal event EventHandler ResetRequested;

		internal event EventHandler SaveAsRequested;

		internal event EventHandler DeleteBlankRowsRequested;

		internal event EventHandler<AccountingPaymentHistoryEntryRequestEventArgs> AddHistoryRequested;

		internal event EventHandler<AccountingPaymentHistoryEntryRequestEventArgs> OutputFutureBalanceRequested;

		internal AccountingPaymentHistoryInputForm ()
		{
			InitializeComponent ();
			AccountingFormButtonAppearanceHelper.Apply (btn発行日を入力, btnリセット, btn今日, btn履歴を入力, btn今後の残高推移を出力, btn別名保存, btn削除);
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

		internal void BindState (AccountingPaymentHistoryFormState state)
		{
			text請求額.Text = ((state == null) ? string.Empty : (state.BillingAmountText ?? string.Empty));
			text実費等.Text = ((state == null) ? string.Empty : (state.ExpenseAmountText ?? string.Empty));
			text源泉処理.Text = ((state == null) ? string.Empty : (state.WithholdingText ?? string.Empty));
			textお預かり金額.Text = ((state == null) ? string.Empty : (state.DepositAmountText ?? string.Empty));
			text領収日.Text = ((state == null) ? string.Empty : (state.ReceiptDateText ?? string.Empty));
			text領収額.Text = ((state == null) ? string.Empty : (state.ReceiptAmountText ?? string.Empty));
		}

		internal void FocusReceiptDate ()
		{
			text領収日.Focus ();
		}

		internal void FocusReceiptAmount ()
		{
			text領収額.Focus ();
		}

		private void BtnIssueDate_Click (object sender, EventArgs e)
		{
			this.IssueDateRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnToday_Click (object sender, EventArgs e)
		{
			this.TodayRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnReset_Click (object sender, EventArgs e)
		{
			this.ResetRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnSaveAs_Click (object sender, EventArgs e)
		{
			this.SaveAsRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnDeleteBlankRows_Click (object sender, EventArgs e)
		{
			this.DeleteBlankRowsRequested?.Invoke (this, EventArgs.Empty);
		}

		private void BtnAddHistory_Click (object sender, EventArgs e)
		{
			this.AddHistoryRequested?.Invoke (this, new AccountingPaymentHistoryEntryRequestEventArgs (CreateRequest ()));
		}

		private void BtnOutputFutureBalance_Click (object sender, EventArgs e)
		{
			this.OutputFutureBalanceRequested?.Invoke (this, new AccountingPaymentHistoryEntryRequestEventArgs (CreateRequest ()));
		}

		private AccountingPaymentHistoryEntryRequest CreateRequest ()
		{
			return new AccountingPaymentHistoryEntryRequest {
				BillingAmountText = text請求額.Text,
				ExpenseAmountText = text実費等.Text,
				WithholdingText = text源泉処理.Text,
				DepositAmountText = textお預かり金額.Text,
				ReceiptDateText = text領収日.Text,
				ReceiptAmountText = text領収額.Text
			};
		}

		private void ShowInvoiceEditRestrictedMessage (object sender, KeyEventArgs e)
		{
			MessageBox.Show (this, "入力フォームでは変更できません。" + Environment.NewLine + "変更は請求書シートで行ってください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
			this.BackColor = System.Drawing.Color.FromArgb (234, 255, 234);
			base.ClientSize = new System.Drawing.Size (525, 605);
			this.Font = new System.Drawing.Font ("Yu Gothic UI", 10f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
			base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.Name = "AccountingPaymentHistoryInputForm";
			base.ShowInTaskbar = false;
			base.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "お支払い履歴入力フォーム";
			this.btn発行日を入力 = new System.Windows.Forms.Button ();
			this.btn発行日を入力.Location = new System.Drawing.Point (64, 20);
			this.btn発行日を入力.Name = "btn発行日を入力";
			this.btn発行日を入力.Size = new System.Drawing.Size (120, 33);
			this.btn発行日を入力.TabIndex = 0;
			this.btn発行日を入力.Text = "発行日を入力";
			this.btn発行日を入力.UseVisualStyleBackColor = true;
			this.btn発行日を入力.Click += new System.EventHandler (BtnIssueDate_Click);
			this.btnリセット = new System.Windows.Forms.Button ();
			this.btnリセット.BackColor = System.Drawing.Color.FromArgb (255, 192, 192);
			this.btnリセット.Location = new System.Drawing.Point (341, 20);
			this.btnリセット.Name = "btnリセット";
			this.btnリセット.Size = new System.Drawing.Size (120, 33);
			this.btnリセット.TabIndex = 3;
			this.btnリセット.Text = "リセット";
			this.btnリセット.UseVisualStyleBackColor = false;
			this.btnリセット.Click += new System.EventHandler (BtnReset_Click);
			this.frame請求書の記載内容 = new CaseInfoSystem.ExcelAddIn.UI.VbaFramePanel ();
			this.frame請求書の記載内容.BackColor = this.BackColor;
			this.frame請求書の記載内容.Caption = "【請求書の記載内容】";
			this.frame請求書の記載内容.Location = new System.Drawing.Point (24, 68);
			this.frame請求書の記載内容.Name = "frame請求書の記載内容";
			this.frame請求書の記載内容.Size = new System.Drawing.Size (475, 150);
			this.frame請求書の記載内容.TabIndex = 4;
			this.lbl請求額 = new System.Windows.Forms.Label ();
			this.lbl請求額.BackColor = System.Drawing.Color.Transparent;
			this.lbl請求額.Location = new System.Drawing.Point (24, 23);
			this.lbl請求額.Name = "lbl請求額";
			this.lbl請求額.Size = new System.Drawing.Size (57, 24);
			this.lbl請求額.TabIndex = 6;
			this.lbl請求額.Text = "請求額";
			this.text請求額 = new System.Windows.Forms.TextBox ();
			this.text請求額.BackColor = this.BackColor;
			this.text請求額.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.text請求額.Location = new System.Drawing.Point (32, 49);
			this.text請求額.Name = "text請求額";
			this.text請求額.ReadOnly = true;
			this.text請求額.Size = new System.Drawing.Size (120, 25);
			this.text請求額.TabIndex = 0;
			this.text請求額.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.text請求額.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			this.lbl請求額円 = new System.Windows.Forms.Label ();
			this.lbl請求額円.BackColor = System.Drawing.Color.Transparent;
			this.lbl請求額円.Location = new System.Drawing.Point (160, 49);
			this.lbl請求額円.Name = "lbl請求額円";
			this.lbl請求額円.Size = new System.Drawing.Size (24, 24);
			this.lbl請求額円.TabIndex = 4;
			this.lbl請求額円.Text = "円";
			this.lbl実費等 = new System.Windows.Forms.Label ();
			this.lbl実費等.BackColor = System.Drawing.Color.Transparent;
			this.lbl実費等.Location = new System.Drawing.Point (196, 23);
			this.lbl実費等.Name = "lbl実費等";
			this.lbl実費等.Size = new System.Drawing.Size (108, 24);
			this.lbl実費等.TabIndex = 5;
			this.lbl実費等.Text = "（うち実費等）";
			this.text実費等 = new System.Windows.Forms.TextBox ();
			this.text実費等.BackColor = this.BackColor;
			this.text実費等.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.text実費等.Location = new System.Drawing.Point (204, 49);
			this.text実費等.Name = "text実費等";
			this.text実費等.ReadOnly = true;
			this.text実費等.Size = new System.Drawing.Size (120, 25);
			this.text実費等.TabIndex = 1;
			this.text実費等.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.text実費等.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			this.lbl実費等円 = new System.Windows.Forms.Label ();
			this.lbl実費等円.BackColor = System.Drawing.Color.Transparent;
			this.lbl実費等円.Location = new System.Drawing.Point (332, 49);
			this.lbl実費等円.Name = "lbl実費等円";
			this.lbl実費等円.Size = new System.Drawing.Size (24, 24);
			this.lbl実費等円.TabIndex = 3;
			this.lbl実費等円.Text = "円";
			this.lbl源泉処理 = new System.Windows.Forms.Label ();
			this.lbl源泉処理.BackColor = System.Drawing.Color.Transparent;
			this.lbl源泉処理.Location = new System.Drawing.Point (380, 23);
			this.lbl源泉処理.Name = "lbl源泉処理";
			this.lbl源泉処理.Size = new System.Drawing.Size (72, 24);
			this.lbl源泉処理.TabIndex = 2;
			this.lbl源泉処理.Text = "源泉処理";
			this.text源泉処理 = new System.Windows.Forms.TextBox ();
			this.text源泉処理.BackColor = this.BackColor;
			this.text源泉処理.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.text源泉処理.Location = new System.Drawing.Point (390, 49);
			this.text源泉処理.Name = "text源泉処理";
			this.text源泉処理.ReadOnly = true;
			this.text源泉処理.Size = new System.Drawing.Size (53, 25);
			this.text源泉処理.TabIndex = 7;
			this.text源泉処理.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
			this.text源泉処理.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			this.lblお預かり金額 = new System.Windows.Forms.Label ();
			this.lblお預かり金額.BackColor = System.Drawing.Color.Transparent;
			this.lblお預かり金額.Location = new System.Drawing.Point (196, 83);
			this.lblお預かり金額.Name = "lblお預かり金額";
			this.lblお預かり金額.Size = new System.Drawing.Size (192, 24);
			this.lblお預かり金額.TabIndex = 9;
			this.lblお預かり金額.Text = "お預かり金額（充当額）";
			this.textお預かり金額 = new System.Windows.Forms.TextBox ();
			this.textお預かり金額.BackColor = this.BackColor;
			this.textお預かり金額.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.textお預かり金額.Location = new System.Drawing.Point (204, 109);
			this.textお預かり金額.Name = "textお預かり金額";
			this.textお預かり金額.ReadOnly = true;
			this.textお預かり金額.Size = new System.Drawing.Size (120, 25);
			this.textお預かり金額.TabIndex = 8;
			this.textお預かり金額.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.textお預かり金額.KeyDown += new System.Windows.Forms.KeyEventHandler (ShowInvoiceEditRestrictedMessage);
			this.lblお預かり金額円 = new System.Windows.Forms.Label ();
			this.lblお預かり金額円.BackColor = System.Drawing.Color.Transparent;
			this.lblお預かり金額円.Location = new System.Drawing.Point (332, 109);
			this.lblお預かり金額円.Name = "lblお預かり金額円";
			this.lblお預かり金額円.Size = new System.Drawing.Size (24, 24);
			this.lblお預かり金額円.TabIndex = 10;
			this.lblお預かり金額円.Text = "円";
			this.frame請求書の記載内容.Controls.Add (this.lbl請求額);
			this.frame請求書の記載内容.Controls.Add (this.text請求額);
			this.frame請求書の記載内容.Controls.Add (this.lbl請求額円);
			this.frame請求書の記載内容.Controls.Add (this.lbl実費等);
			this.frame請求書の記載内容.Controls.Add (this.text実費等);
			this.frame請求書の記載内容.Controls.Add (this.lbl実費等円);
			this.frame請求書の記載内容.Controls.Add (this.lbl源泉処理);
			this.frame請求書の記載内容.Controls.Add (this.text源泉処理);
			this.frame請求書の記載内容.Controls.Add (this.lblお預かり金額);
			this.frame請求書の記載内容.Controls.Add (this.textお預かり金額);
			this.frame請求書の記載内容.Controls.Add (this.lblお預かり金額円);
			this.frame領収内容 = new CaseInfoSystem.ExcelAddIn.UI.VbaFramePanel ();
			this.frame領収内容.BackColor = this.BackColor;
			this.frame領収内容.Caption = "【領収内容】";
			this.frame領収内容.Location = new System.Drawing.Point (24, 230);
			this.frame領収内容.Name = "frame領収内容";
			this.frame領収内容.Size = new System.Drawing.Size (475, 156);
			this.frame領収内容.TabIndex = 1;
			this.lbl領収日 = new System.Windows.Forms.Label ();
			this.lbl領収日.BackColor = System.Drawing.Color.Transparent;
			this.lbl領収日.Location = new System.Drawing.Point (24, 21);
			this.lbl領収日.Name = "lbl領収日";
			this.lbl領収日.Size = new System.Drawing.Size (48, 24);
			this.lbl領収日.TabIndex = 5;
			this.lbl領収日.Text = "領収日";
			this.text領収日 = new System.Windows.Forms.TextBox ();
			this.text領収日.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.text領収日.Location = new System.Drawing.Point (32, 47);
			this.text領収日.Name = "text領収日";
			this.text領収日.Size = new System.Drawing.Size (120, 25);
			this.text領収日.TabIndex = 0;
			this.text領収日.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.btn今日 = new System.Windows.Forms.Button ();
			this.btn今日.Location = new System.Drawing.Point (160, 47);
			this.btn今日.Name = "btn今日";
			this.btn今日.Size = new System.Drawing.Size (52, 25);
			this.btn今日.TabIndex = 2;
			this.btn今日.Text = "今日";
			this.btn今日.UseVisualStyleBackColor = true;
			this.btn今日.Click += new System.EventHandler (BtnToday_Click);
			this.lbl領収額 = new System.Windows.Forms.Label ();
			this.lbl領収額.BackColor = System.Drawing.Color.Transparent;
			this.lbl領収額.Location = new System.Drawing.Point (24, 86);
			this.lbl領収額.Name = "lbl領収額";
			this.lbl領収額.Size = new System.Drawing.Size (48, 24);
			this.lbl領収額.TabIndex = 6;
			this.lbl領収額.Text = "領収額";
			this.text領収額 = new System.Windows.Forms.TextBox ();
			this.text領収額.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.text領収額.Location = new System.Drawing.Point (32, 112);
			this.text領収額.Name = "text領収額";
			this.text領収額.Size = new System.Drawing.Size (120, 25);
			this.text領収額.TabIndex = 1;
			this.text領収額.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.text領収額.KeyPress += new System.Windows.Forms.KeyPressEventHandler (NumericTextBox_KeyPress);
			this.text領収額.Leave += new System.EventHandler (CurrencyTextBox_Leave);
			this.lbl領収額円 = new System.Windows.Forms.Label ();
			this.lbl領収額円.BackColor = System.Drawing.Color.Transparent;
			this.lbl領収額円.Location = new System.Drawing.Point (160, 112);
			this.lbl領収額円.Name = "lbl領収額円";
			this.lbl領収額円.Size = new System.Drawing.Size (24, 24);
			this.lbl領収額円.TabIndex = 7;
			this.lbl領収額円.Text = "円";
			this.btn履歴を入力 = new System.Windows.Forms.Button ();
			this.btn履歴を入力.BackColor = System.Drawing.Color.FromArgb (0, 0, 128);
			this.btn履歴を入力.ForeColor = System.Drawing.Color.White;
			this.btn履歴を入力.Location = new System.Drawing.Point (251, 29);
			this.btn履歴を入力.Name = "btn履歴を入力";
			this.btn履歴を入力.Size = new System.Drawing.Size (181, 52);
			this.btn履歴を入力.TabIndex = 3;
			this.btn履歴を入力.Text = "この領収内容で\r\n履歴(1回分)を入力";
			this.btn履歴を入力.UseVisualStyleBackColor = false;
			this.btn履歴を入力.Click += new System.EventHandler (BtnAddHistory_Click);
			this.btn今後の残高推移を出力 = new System.Windows.Forms.Button ();
			this.btn今後の残高推移を出力.Location = new System.Drawing.Point (251, 96);
			this.btn今後の残高推移を出力.Name = "btn今後の残高推移を出力";
			this.btn今後の残高推移を出力.Size = new System.Drawing.Size (178, 31);
			this.btn今後の残高推移を出力.TabIndex = 4;
			this.btn今後の残高推移を出力.Text = "今後の残高推移を出力";
			this.btn今後の残高推移を出力.UseVisualStyleBackColor = true;
			this.btn今後の残高推移を出力.Click += new System.EventHandler (BtnOutputFutureBalance_Click);
			this.frame領収内容.Controls.Add (this.lbl領収日);
			this.frame領収内容.Controls.Add (this.text領収日);
			this.frame領収内容.Controls.Add (this.btn今日);
			this.frame領収内容.Controls.Add (this.lbl領収額);
			this.frame領収内容.Controls.Add (this.text領収額);
			this.frame領収内容.Controls.Add (this.lbl領収額円);
			this.frame領収内容.Controls.Add (this.btn履歴を入力);
			this.frame領収内容.Controls.Add (this.btn今後の残高推移を出力);
			this.btn別名保存 = new System.Windows.Forms.Button ();
			this.btn別名保存.Location = new System.Drawing.Point (64, 394);
			this.btn別名保存.Name = "btn別名保存";
			this.btn別名保存.Size = new System.Drawing.Size (120, 33);
			this.btn別名保存.TabIndex = 2;
			this.btn別名保存.Text = "別名保存";
			this.btn別名保存.UseVisualStyleBackColor = true;
			this.btn別名保存.Click += new System.EventHandler (BtnSaveAs_Click);
			this.lbl修正削除見出し = new System.Windows.Forms.Label ();
			this.lbl修正削除見出し.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.lbl修正削除見出し.Location = new System.Drawing.Point (0, 440);
			this.lbl修正削除見出し.Name = "lbl修正削除見出し";
			this.lbl修正削除見出し.Size = new System.Drawing.Size (525, 28);
			this.lbl修正削除見出し.TabIndex = 6;
			this.lbl修正削除見出し.Text = "------《修正・削除の手順》------------------------------------------------------------";
			this.lbl手順1 = new System.Windows.Forms.Label ();
			this.lbl手順1.BackColor = System.Drawing.Color.Transparent;
			this.lbl手順1.Location = new System.Drawing.Point (24, 474);
			this.lbl手順1.Name = "lbl手順1";
			this.lbl手順1.Size = new System.Drawing.Size (470, 24);
			this.lbl手順1.TabIndex = 7;
			this.lbl手順1.Text = "⑴\u3000履歴シート上の修正・削除したい回の領収日を手動でクリア";
			this.lbl手順2 = new System.Windows.Forms.Label ();
			this.lbl手順2.BackColor = System.Drawing.Color.Transparent;
			this.lbl手順2.Location = new System.Drawing.Point (24, 498);
			this.lbl手順2.Name = "lbl手順2";
			this.lbl手順2.Size = new System.Drawing.Size (470, 24);
			this.lbl手順2.TabIndex = 8;
			this.lbl手順2.Text = "⑵\u3000削除ボタン↓をクリック（履歴シート上の領収日が空白の回を削除します）";
			this.btn削除 = new System.Windows.Forms.Button ();
			this.btn削除.Font = new System.Drawing.Font ("Yu Gothic UI", 8f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128);
			this.btn削除.Location = new System.Drawing.Point (64, 521);
			this.btn削除.Name = "btn削除";
			this.btn削除.Size = new System.Drawing.Size (120, 33);
			this.btn削除.TabIndex = 9;
			this.btn削除.Text = "領収日空白回の削除";
			this.btn削除.UseVisualStyleBackColor = true;
			this.btn削除.Click += new System.EventHandler (BtnDeleteBlankRows_Click);
			this.lbl手順3 = new System.Windows.Forms.Label ();
			this.lbl手順3.BackColor = System.Drawing.Color.Transparent;
			this.lbl手順3.Location = new System.Drawing.Point (24, 559);
			this.lbl手順3.Name = "lbl手順3";
			this.lbl手順3.Size = new System.Drawing.Size (470, 24);
			this.lbl手順3.TabIndex = 10;
			this.lbl手順3.Text = "⑶\u3000正しい領収日・領収額で再入力（履歴シートは日付順にソートされます）";
			base.Controls.Add (this.btn発行日を入力);
			base.Controls.Add (this.btnリセット);
			base.Controls.Add (this.frame請求書の記載内容);
			base.Controls.Add (this.frame領収内容);
			base.Controls.Add (this.btn別名保存);
			base.Controls.Add (this.lbl修正削除見出し);
			base.Controls.Add (this.lbl手順1);
			base.Controls.Add (this.lbl手順2);
			base.Controls.Add (this.btn削除);
			base.Controls.Add (this.lbl手順3);
			base.ResumeLayout (false);
		}
	}
}

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingPaymentHistoryEntryRequestEventArgs : EventArgs
	{
		internal AccountingPaymentHistoryEntryRequest Request { get; }

		internal AccountingPaymentHistoryEntryRequestEventArgs (AccountingPaymentHistoryEntryRequest request)
		{
			Request = request ?? new AccountingPaymentHistoryEntryRequest ();
		}
	}
}
