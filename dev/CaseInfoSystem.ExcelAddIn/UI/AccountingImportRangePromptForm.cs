using System;
using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingImportRangePromptForm : Form
	{
		private readonly TextBox _txtStartRound;

		private readonly TextBox _txtEndRound;

		private readonly Button _btnConfirm;

		private readonly Button _btnExcelClose;

		private int StartRound => ParseRound (_txtStartRound.Text);

		private int EndRound => ParseRound (_txtEndRound.Text);

		internal event EventHandler<AccountingImportRangePromptConfirmedEventArgs> Confirmed;

		internal event EventHandler ExcelCloseRequested;

		internal AccountingImportRangePromptForm (int initialStartRound, int initialEndRound)
		{
			SuspendLayout ();
			base.AutoScaleMode = AutoScaleMode.None;
			BackColor = SystemColors.Menu;
			base.ClientSize = new Size (508, 320);
			Font = new Font ("Yu Gothic UI", 10f, FontStyle.Regular, GraphicsUnit.Point, 128);
			base.FormBorderStyle = FormBorderStyle.FixedDialog;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.ShowInTaskbar = false;
			base.StartPosition = FormStartPosition.CenterParent;
			_btnExcelClose = new Button {
				Location = new Point (364, 10),
				Name = "btnExcelClose",
				Size = new Size (120, 32),
				TabIndex = 4,
				Text = "Excelを閉じる",
				UseVisualStyleBackColor = true
			};
			_btnExcelClose.Click += BtnExcelClose_Click;
			Text = "お支払い履歴から取り込む";
			Label value = new Label {
				Location = new Point (16, 54),
				Name = "lblGuide",
				Size = new Size (478, 24),
				TabIndex = 6,
				Text = "お支払い履歴で指定した範囲の合計額を会計依頼書の対応する欄に入力します"
			};
			Label value2 = new Label {
				Location = new Point (0, 92),
				Name = "lblDivider",
				Size = new Size (508, 18),
				TabIndex = 7,
				Text = "------《作業手順》--------------------------------------------------------------"
			};
			Label value3 = new Label {
				Location = new Point (16, 126),
				Name = "lblStep1",
				Size = new Size (160, 22),
				TabIndex = 8,
				Text = "⑴\u3000対象範囲を指定"
			};
			_txtStartRound = new TextBox {
				BorderStyle = BorderStyle.FixedSingle,
				Location = new Point (45, 158),
				Name = "txtStartRound",
				Size = new Size (28, 25),
				TabIndex = 0,
				Text = ((initialStartRound > 0) ? initialStartRound.ToString () : string.Empty),
				TextAlign = HorizontalAlignment.Center
			};
			_txtStartRound.KeyPress += RoundTextBox_KeyPress;
			Label value4 = new Label {
				Location = new Point (79, 160),
				Name = "lblStartRoundSuffix",
				Size = new Size (58, 22),
				TabIndex = 9,
				Text = "回目から"
			};
			_txtEndRound = new TextBox {
				BorderStyle = BorderStyle.FixedSingle,
				Location = new Point (142, 158),
				Name = "txtEndRound",
				Size = new Size (28, 25),
				TabIndex = 1,
				Text = ((initialEndRound > 0) ? initialEndRound.ToString () : string.Empty),
				TextAlign = HorizontalAlignment.Center
			};
			_txtEndRound.KeyPress += RoundTextBox_KeyPress;
			Label value5 = new Label {
				Location = new Point (176, 160),
				Name = "lblEndRoundSuffix",
				Size = new Size (304, 22),
				TabIndex = 10,
				Text = "回目の支払い（範囲を絞るときは修正してください）"
			};
			Label value6 = new Label {
				Location = new Point (16, 196),
				Name = "lblStep2",
				Size = new Size (472, 22),
				TabIndex = 11,
				Text = "⑵\u3000税処理前の金額を表示させる場所を指定（黄色エリア内のセルを1つ選択）"
			};
			Label value7 = new Label {
				Location = new Point (16, 230),
				Name = "lblStep3",
				Size = new Size (410, 22),
				TabIndex = 12,
				Text = "⑶\u3000決定ボタンをクリック（費用項目は適宜手入力してください）"
			};
			_btnConfirm = new Button {
				BackColor = Color.PowderBlue,
				ForeColor = Color.Black,
				Location = new Point (56, 266),
				Name = "btnConfirm",
				Size = new Size (108, 34),
				TabIndex = 2,
				Text = "決\u3000\u3000定",
				UseVisualStyleBackColor = false
			};
			_btnConfirm.Click += BtnConfirm_Click;
			base.Controls.Add (value);
			base.Controls.Add (value2);
			base.Controls.Add (value3);
			base.Controls.Add (_txtStartRound);
			base.Controls.Add (value4);
			base.Controls.Add (_txtEndRound);
			base.Controls.Add (value5);
			base.Controls.Add (value6);
			base.Controls.Add (value7);
			base.Controls.Add (_btnConfirm);
			base.Controls.Add (_btnExcelClose);
			AccountingFormButtonAppearanceHelper.Apply (_btnConfirm, _btnExcelClose);
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

		internal void ClearRequestHandlers ()
		{
			Confirmed = null;
			ExcelCloseRequested = null;
		}

		private void BtnConfirm_Click (object sender, EventArgs e)
		{
			if (StartRound <= 0 || EndRound <= 0) {
				MessageBox.Show (this, "対象範囲を数字で指定してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			} else if (StartRound > 60 || EndRound > 60) {
				MessageBox.Show (this, "60回目までの範囲を指定してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			} else if (StartRound > EndRound) {
				MessageBox.Show (this, "終期は始期以上で指定してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			} else {
				this.Confirmed?.Invoke (this, new AccountingImportRangePromptConfirmedEventArgs (new AccountingImportRange (StartRound, EndRound)));
			}
		}

		private void BtnExcelClose_Click (object sender, EventArgs e)
		{
			this.ExcelCloseRequested?.Invoke (this, EventArgs.Empty);
		}

		protected override void Dispose (bool disposing)
		{
			if (disposing) {
				ClearRequestHandlers ();
			}
			base.Dispose (disposing);
		}

		private static void RoundTextBox_KeyPress (object sender, KeyPressEventArgs e)
		{
			if (!char.IsControl (e.KeyChar) && !char.IsDigit (e.KeyChar)) {
				e.Handled = true;
			}
		}

		private static int ParseRound (string text)
		{
			int result;
			return int.TryParse ((text ?? string.Empty).Trim (), out result) ? result : 0;
		}
	}
}

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingImportRangePromptConfirmedEventArgs : EventArgs
	{
		internal AccountingImportRange ImportRange { get; }

		internal AccountingImportRangePromptConfirmedEventArgs (AccountingImportRange importRange)
		{
			ImportRange = importRange;
		}
	}
}
