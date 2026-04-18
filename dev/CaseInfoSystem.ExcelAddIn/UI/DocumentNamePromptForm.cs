using System;
using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class DocumentNamePromptForm : Form
	{
		private readonly TextBox _txtDocumentName;

		private readonly Button _btnOk;

		private readonly Button _btnCancel;

		internal string ResultDocumentName => (_txtDocumentName.Text ?? string.Empty).Trim ();

		private DocumentNamePromptForm (string initialDocumentName)
		{
			Text = "文書名確認";
			Font = new Font ("Yu Gothic UI", 10f, FontStyle.Regular, GraphicsUnit.Point, 128);
			base.FormBorderStyle = FormBorderStyle.FixedDialog;
			base.StartPosition = FormStartPosition.CenterParent;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.ShowInTaskbar = false;
			base.ClientSize = new Size (520, 156);
			Label value = new Label {
				AutoSize = true,
				Left = 18,
				Top = 18,
				Text = "作成する文書名を確認してください。"
			};
			Label value2 = new Label {
				AutoSize = true,
				Left = 18,
				Top = 56,
				Text = "文書名"
			};
			_txtDocumentName = new TextBox {
				Left = 18,
				Top = 80,
				Width = 484,
				Text = (initialDocumentName ?? string.Empty)
			};
			_txtDocumentName.SelectAll ();
			_btnOk = new Button {
				Left = 322,
				Top = 116,
				Width = 86,
				Text = "OK",
				DialogResult = DialogResult.OK
			};
			_btnOk.Click += BtnOk_Click;
			_btnCancel = new Button {
				Left = 416,
				Top = 116,
				Width = 86,
				Text = "キャンセル",
				DialogResult = DialogResult.Cancel
			};
			base.AcceptButton = _btnOk;
			base.CancelButton = _btnCancel;
			base.Controls.Add (value);
			base.Controls.Add (value2);
			base.Controls.Add (_txtDocumentName);
			base.Controls.Add (_btnOk);
			base.Controls.Add (_btnCancel);
		}

		internal static bool TryPrompt (IWin32Window owner, string initialDocumentName, out string finalDocumentName)
		{
			using (DocumentNamePromptForm documentNamePromptForm = new DocumentNamePromptForm (initialDocumentName)) {
				DialogResult dialogResult = ((owner == null) ? documentNamePromptForm.ShowDialog () : documentNamePromptForm.ShowDialog (owner));
				if (dialogResult != DialogResult.OK) {
					finalDocumentName = string.Empty;
					return false;
				}
				finalDocumentName = documentNamePromptForm.ResultDocumentName;
				return finalDocumentName.Length > 0;
			}
		}

		private void BtnOk_Click (object sender, EventArgs e)
		{
			if (ResultDocumentName.Length <= 0) {
				MessageBox.Show (this, "文書名を入力してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				base.DialogResult = DialogResult.None;
				_txtDocumentName.Focus ();
			}
		}
	}
}
