using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingReverseGoalSeekForm : Form
	{
		private readonly TextBox _txtTargetAmount;

		private readonly Button _btnCalculate;

		private readonly Button _btnClose;

		private bool _allowCloseByButton;

		internal event EventHandler<AccountingReverseGoalSeekConfirmedEventArgs> Confirmed;

		internal event EventHandler Canceled;

		internal AccountingReverseGoalSeekForm ()
		{
			SuspendLayout ();
			base.AutoScaleMode = AutoScaleMode.None;
			BackColor = Color.White;
			base.ClientSize = new Size (474, 374);
			Font = new Font ("Yu Gothic UI", 10f, FontStyle.Regular, GraphicsUnit.Point, 128);
			base.FormBorderStyle = FormBorderStyle.FixedDialog;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.ShowInTaskbar = false;
			base.StartPosition = FormStartPosition.CenterParent;
			Text = "逆算ツール";
			Label value = new Label {
				Location = new Point (14, 14),
				Name = "lblGuide",
				Size = new Size (438, 40),
				TabIndex = 5,
				Text = "・丸めた数字で請求したい → 請求額から値引き額を逆算できます\r\n・丸めた金額を受け取った → 領収額から税処理前の価額を逆算できます"
			};
			Label value2 = new Label {
				Location = new Point (0, 64),
				Name = "lblDivider",
				Size = new Size (474, 22),
				TabIndex = 6,
				Text = "----《作業手順》--------------------------------------------------------------------------"
			};
			Label value3 = new Label {
				Location = new Point (18, 95),
				Name = "lblStep1",
				Size = new Size (418, 22),
				TabIndex = 7,
				Text = "⑴  請求額（or 領収額）を幾らにしたいですか？ （金額を入力）"
			};
			_txtTargetAmount = new TextBox {
				BorderStyle = BorderStyle.FixedSingle,
				Location = new Point (56, 133),
				Name = "txtTargetAmount",
				Size = new Size (106, 25),
				TabIndex = 0,
				TextAlign = HorizontalAlignment.Right
			};
			_txtTargetAmount.TextChanged += TxtTargetAmount_TextChanged;
			Label value4 = new Label {
				Location = new Point (169, 133),
				Name = "lblYen",
				Size = new Size (24, 24),
				TabIndex = 8,
				Text = "円"
			};
			Label value5 = new Label {
				Location = new Point (18, 180),
				Name = "lblStep2",
				Size = new Size (434, 22),
				TabIndex = 9,
				Text = "⑵  値引き額等を表示させる場所を指定（黄色エリア内のセルを1つ選択）"
			};
			Label value6 = new Label {
				Location = new Point (18, 220),
				Name = "lblStep3",
				Size = new Size (210, 22),
				TabIndex = 10,
				Text = "⑶  逆算ボタンをクリック"
			};
			Label value7 = new Label {
				Location = new Point (38, 246),
				Name = "lblResultNote",
				Size = new Size (332, 48),
				TabIndex = 11,
				Text = "→ ⑵で選択したセルに計算結果が表示されます\r\n   それと連動して “④請求金額” が⑴の金額になります"
			};
			_btnCalculate = new Button {
				BackColor = Color.PowderBlue,
				ForeColor = Color.Black,
				Location = new Point (56, 314),
				Name = "btnCalculate",
				Size = new Size (107, 40),
				TabIndex = 1,
				Text = "逆\u3000算",
				UseVisualStyleBackColor = false
			};
			_btnCalculate.Click += BtnCalculate_Click;
			_btnClose = new Button {
				Location = new Point (296, 314),
				Name = "btnClose",
				Size = new Size (107, 40),
				TabIndex = 2,
				Text = "閉じる",
				UseVisualStyleBackColor = true
			};
			_btnClose.Click += BtnClose_Click;
			base.Controls.Add (value);
			base.Controls.Add (value2);
			base.Controls.Add (value3);
			base.Controls.Add (_txtTargetAmount);
			base.Controls.Add (value4);
			base.Controls.Add (value5);
			base.Controls.Add (value6);
			base.Controls.Add (value7);
			base.Controls.Add (_btnCalculate);
			base.Controls.Add (_btnClose);
			AccountingFormButtonAppearanceHelper.Apply (_btnCalculate, _btnClose);
			ButtonCursorHelper.ApplyHandCursor (this);
			ResumeLayout (performLayout: false);
		}

		internal void ShowModeless (IWin32Window owner)
		{
			if (owner == null) {
				Show ();
			} else {
				Show (owner);
			}
		}

		internal void CloseByCode ()
		{
			_allowCloseByButton = true;
			Close ();
		}

		private void BtnCalculate_Click (object sender, EventArgs e)
		{
			if (!TryParseTargetAmount (out var targetAmount)) {
				MessageBox.Show (this, "目標金額を数字で入力してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			} else {
				this.Confirmed?.Invoke (this, new AccountingReverseGoalSeekConfirmedEventArgs (new AccountingReverseGoalSeekRequest (targetAmount)));
			}
		}

		private void BtnClose_Click (object sender, EventArgs e)
		{
			CloseByCode ();
		}

		private void TxtTargetAmount_TextChanged (object sender, EventArgs e)
		{
			int selectionStart = _txtTargetAmount.SelectionStart;
			string text = (_txtTargetAmount.Text ?? string.Empty).Replace (",", string.Empty);
			if (double.TryParse (text, NumberStyles.Number, CultureInfo.InvariantCulture, out var result)) {
				string text2 = result.ToString ("#,##0", CultureInfo.InvariantCulture);
				if (!string.Equals (text2, _txtTargetAmount.Text, StringComparison.Ordinal)) {
					_txtTargetAmount.TextChanged -= TxtTargetAmount_TextChanged;
					_txtTargetAmount.Text = text2;
					_txtTargetAmount.SelectionStart = Math.Min (_txtTargetAmount.Text.Length, selectionStart + (text2.Length - text.Length));
					_txtTargetAmount.TextChanged += TxtTargetAmount_TextChanged;
				}
			}
		}

		protected override void OnFormClosing (FormClosingEventArgs e)
		{
			if (e.CloseReason == CloseReason.UserClosing && !_allowCloseByButton) {
				MessageBox.Show (this, "ボタンで閉じてください", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				e.Cancel = true;
			} else {
				base.OnFormClosing (e);
			}
		}

		protected override void OnFormClosed (FormClosedEventArgs e)
		{
			base.OnFormClosed (e);
			this.Canceled?.Invoke (this, EventArgs.Empty);
		}

		private bool TryParseTargetAmount (out double targetAmount)
		{
			string s = (_txtTargetAmount.Text ?? string.Empty).Replace (",", string.Empty).Trim ();
			return double.TryParse (s, NumberStyles.Number, CultureInfo.InvariantCulture, out targetAmount);
		}
	}
}

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingReverseGoalSeekConfirmedEventArgs : EventArgs
	{
		internal AccountingReverseGoalSeekRequest Request { get; }

		internal AccountingReverseGoalSeekConfirmedEventArgs (AccountingReverseGoalSeekRequest request)
		{
			Request = request;
		}
	}
}
