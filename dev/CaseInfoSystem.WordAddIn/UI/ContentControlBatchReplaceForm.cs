using System;
using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.WordAddIn.Services;

namespace CaseInfoSystem.WordAddIn.UI
{
	internal sealed class ContentControlBatchReplaceForm : Form
	{
		private readonly TextBox _oldTagTextBox;

		private readonly TextBox _newTagTextBox;

		private readonly TextBox _oldTitleTextBox;

		private readonly TextBox _newTitleTextBox;

		private readonly CheckBox _partialMatchCheckBox;

		private readonly Label _statusLabel;

		private readonly Button _replaceNextButton;

		private readonly Button _executeButton;

		private readonly Button _cancelButton;

		public ContentControlBatchReplaceService.ReplaceRequest ReplaceRequest { get; private set; }

		public ContentControlBatchReplaceForm ()
		{
			Font = new Font ("Yu Gothic UI", 9f, FontStyle.Regular, GraphicsUnit.Point, 128);
			base.FormBorderStyle = FormBorderStyle.FixedDialog;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.ShowInTaskbar = false;
			base.StartPosition = FormStartPosition.CenterParent;
			Text = "コンテンツコントロール一括置換";
			base.ClientSize = new Size (460, 295);
			Label value = new Label {
				AutoSize = false,
				Location = new Point (20, 15),
				Size = new Size (420, 34),
				Text = "アクティブ文書内のテキスト / リッチテキストのコンテンツコントロールを対象に、Title と Tag を一括置換します。"
			};
			Label value2 = CreateLabel ("旧タグ", 20, 60);
			_oldTagTextBox = CreateTextBox (130, 56);
			_oldTagTextBox.Enter += TextBox_Enter;
			Label value3 = CreateLabel ("新タグ", 20, 95);
			_newTagTextBox = CreateTextBox (130, 91);
			_newTagTextBox.Enter += TextBox_Enter;
			Label value4 = CreateLabel ("旧タイトル", 20, 130);
			_oldTitleTextBox = CreateTextBox (130, 126);
			_oldTitleTextBox.Enter += TextBox_Enter;
			Label value5 = CreateLabel ("新タイトル", 20, 165);
			_newTitleTextBox = CreateTextBox (130, 161);
			_newTitleTextBox.Enter += TextBox_Enter;
			_partialMatchCheckBox = new CheckBox {
				AutoSize = true,
				Location = new Point (20, 198),
				Text = "部分一致で置換する"
			};
			_statusLabel = new Label {
				AutoEllipsis = true,
				BorderStyle = BorderStyle.FixedSingle,
				Location = new Point (20, 224),
				Size = new Size (420, 28),
				TextAlign = ContentAlignment.MiddleLeft
			};
			_replaceNextButton = new Button {
				Text = "下へ1件置換",
				Location = new Point (165, 260),
				Size = new Size (90, 28)
			};
			_replaceNextButton.Click += ReplaceNextButton_Click;
			_executeButton = new Button {
				Text = "一括置換",
				Location = new Point (260, 260),
				Size = new Size (85, 28)
			};
			_executeButton.Click += ExecuteButton_Click;
			_cancelButton = new Button {
				Text = "閉じる",
				Location = new Point (355, 260),
				Size = new Size (85, 28),
				DialogResult = DialogResult.Cancel
			};
			base.AcceptButton = _replaceNextButton;
			base.CancelButton = _cancelButton;
			base.Controls.Add (value);
			base.Controls.Add (value2);
			base.Controls.Add (_oldTagTextBox);
			base.Controls.Add (value3);
			base.Controls.Add (_newTagTextBox);
			base.Controls.Add (value4);
			base.Controls.Add (_oldTitleTextBox);
			base.Controls.Add (value5);
			base.Controls.Add (_newTitleTextBox);
			base.Controls.Add (_partialMatchCheckBox);
			base.Controls.Add (_statusLabel);
			base.Controls.Add (_replaceNextButton);
			base.Controls.Add (_executeButton);
			base.Controls.Add (_cancelButton);
		}

		private void ExecuteButton_Click (object sender, EventArgs e)
		{
			ContentControlBatchReplaceService.ReplaceRequest replaceRequest = BuildReplaceRequest ();
			if (!ContentControlBatchReplaceService.HasAnyTarget (replaceRequest)) {
				MessageBox.Show (this, "旧タグまたは旧タイトルのどちらかを入力してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return;
			}
			ReplaceRequest = replaceRequest;
			base.DialogResult = DialogResult.OK;
			Close ();
		}

		private void ReplaceNextButton_Click (object sender, EventArgs e)
		{
			ContentControlBatchReplaceService.ReplaceRequest request = BuildReplaceRequest ();
			if (!ContentControlBatchReplaceService.HasAnyTarget (request)) {
				MessageBox.Show (this, "旧タグまたは旧タイトルのどちらかを入力してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return;
			}
			try {
				ContentControlBatchReplaceService.NextReplaceResult result = Globals.ThisAddIn.ReplaceNextContentControlFromSelection (request);
				_statusLabel.Text = ContentControlBatchReplaceService.BuildNextReplaceMessage (result);
			} catch (Exception ex) {
				MessageBox.Show (this, ex.Message, "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}

		private ContentControlBatchReplaceService.ReplaceRequest BuildReplaceRequest ()
		{
			return new ContentControlBatchReplaceService.ReplaceRequest {
				OldTag = _oldTagTextBox.Text,
				NewTag = _newTagTextBox.Text,
				OldTitle = _oldTitleTextBox.Text,
				NewTitle = _newTitleTextBox.Text,
				UsePartialMatch = _partialMatchCheckBox.Checked
			};
		}

		private static void TextBox_Enter (object sender, EventArgs e)
		{
			if (sender is TextBox textBox) {
				textBox.ImeMode = ImeMode.On;
			}
		}

		private static Label CreateLabel (string text, int x, int y)
		{
			return new Label {
				AutoSize = true,
				Location = new Point (x, y + 4),
				Text = text
			};
		}

		private static TextBox CreateTextBox (int x, int y)
		{
			return new TextBox {
				Location = new Point (x, y),
				Size = new Size (310, 23)
			};
		}
	}
}
