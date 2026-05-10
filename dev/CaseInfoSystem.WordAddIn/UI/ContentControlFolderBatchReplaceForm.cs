using System;
using System.IO;
using System.Windows.Forms;
using CaseInfoSystem.WordAddIn.Services;

namespace CaseInfoSystem.WordAddIn.UI
{
	internal sealed class ContentControlFolderBatchReplaceForm : Form
	{
		private readonly TextBox _folderTextBox;

		private readonly Button _browseButton;

		private readonly TextBox _oldTagTextBox;

		private readonly TextBox _newTagTextBox;

		private readonly TextBox _oldTitleTextBox;

		private readonly TextBox _newTitleTextBox;

		private readonly TextBox _oldDisplayTextBox;

		private readonly TextBox _newDisplayTextBox;

		private readonly CheckBox _partialMatchCheckBox;

		private readonly CheckBox _backupCheckBox;

		private readonly Label _statusLabel;

		public ContentControlFolderBatchReplaceService.FolderReplaceRequest ReplaceRequest { get; private set; }

		internal ContentControlFolderBatchReplaceForm (string initialDirectory)
		{
			Text = "雛形フォルダ CC 一括置換";
			Width = 560;
			Height = 400;
			FormBorderStyle = FormBorderStyle.FixedDialog;
			MaximizeBox = false;
			MinimizeBox = false;
			StartPosition = FormStartPosition.CenterParent;

			Label descriptionLabel = new Label {
				Left = 16,
				Top = 14,
				Width = 512,
				Height = 36,
				Text = "雛形フォルダ直下の Word 文書を対象に、テキスト / リッチテキストのコンテンツコントロールの Title、Tag、表示文字を置換します。"
			};
			base.Controls.Add (descriptionLabel);

			base.Controls.Add (CreateLabel ("雛形フォルダ", 16, 59));
			_folderTextBox = CreateTextBox (114, 56, 334);
			_folderTextBox.Text = initialDirectory ?? string.Empty;
			base.Controls.Add (_folderTextBox);

			_browseButton = new Button {
				Left = 456,
				Top = 54,
				Width = 72,
				Height = 26,
				Text = "参照..."
			};
			_browseButton.Click += BrowseButton_Click;
			base.Controls.Add (_browseButton);

			base.Controls.Add (CreateLabel ("置換前 Tag", 16, 94));
			_oldTagTextBox = CreateTextBox (114, 91, 414);
			base.Controls.Add (_oldTagTextBox);

			base.Controls.Add (CreateLabel ("置換後 Tag", 16, 129));
			_newTagTextBox = CreateTextBox (114, 126, 414);
			base.Controls.Add (_newTagTextBox);

			base.Controls.Add (CreateLabel ("置換前 Title", 16, 164));
			_oldTitleTextBox = CreateTextBox (114, 161, 414);
			base.Controls.Add (_oldTitleTextBox);

			base.Controls.Add (CreateLabel ("置換後 Title", 16, 199));
			_newTitleTextBox = CreateTextBox (114, 196, 414);
			base.Controls.Add (_newTitleTextBox);

			base.Controls.Add (CreateLabel ("置換前 表示", 16, 234));
			_oldDisplayTextBox = CreateTextBox (114, 231, 414);
			base.Controls.Add (_oldDisplayTextBox);

			base.Controls.Add (CreateLabel ("置換後 表示", 16, 269));
			_newDisplayTextBox = CreateTextBox (114, 266, 414);
			base.Controls.Add (_newDisplayTextBox);

			_partialMatchCheckBox = new CheckBox {
				Left = 114,
				Top = 299,
				Width = 160,
				Height = 24,
				Text = "部分一致で置換する"
			};
			base.Controls.Add (_partialMatchCheckBox);

			_backupCheckBox = new CheckBox {
				Left = 282,
				Top = 299,
				Width = 180,
				Height = 24,
				Text = "変更前バックアップを作成",
				Checked = true
			};
			base.Controls.Add (_backupCheckBox);

			_statusLabel = new Label {
				Left = 16,
				Top = 327,
				Width = 330,
				Height = 24
			};
			base.Controls.Add (_statusLabel);

			Button okButton = new Button {
				Left = 352,
				Top = 323,
				Width = 86,
				Height = 30,
				Text = "一括置換"
			};
			okButton.Click += OkButton_Click;
			base.Controls.Add (okButton);

			Button cancelButton = new Button {
				Left = 446,
				Top = 323,
				Width = 82,
				Height = 30,
				Text = "キャンセル",
				DialogResult = DialogResult.Cancel
			};
			base.Controls.Add (cancelButton);

			base.AcceptButton = okButton;
			base.CancelButton = cancelButton;
		}

		private void BrowseButton_Click (object sender, EventArgs e)
		{
			using (FolderBrowserDialog dialog = new FolderBrowserDialog ()) {
				dialog.Description = "置換対象の雛形フォルダを選択してください。";
				dialog.ShowNewFolderButton = false;
				string currentDirectory = (_folderTextBox.Text ?? string.Empty).Trim ();
				if (Directory.Exists (currentDirectory)) {
					dialog.SelectedPath = currentDirectory;
				}
				if (dialog.ShowDialog (this) == DialogResult.OK) {
					_folderTextBox.Text = dialog.SelectedPath ?? string.Empty;
				}
			}
		}

		private void OkButton_Click (object sender, EventArgs e)
		{
			ContentControlFolderBatchReplaceService.FolderReplaceRequest request = BuildReplaceRequest ();
			if (string.IsNullOrWhiteSpace (request.TemplateDirectory) || !Directory.Exists (request.TemplateDirectory)) {
				_statusLabel.Text = "雛形フォルダを選択してください。";
				return;
			}
			if (!ContentControlFolderBatchReplaceService.HasAnyTarget (request)) {
				_statusLabel.Text = "置換前 Tag / Title / 表示文字のいずれかを入力してください。";
				return;
			}

			ReplaceRequest = request;
			base.DialogResult = DialogResult.OK;
			Close ();
		}

		private ContentControlFolderBatchReplaceService.FolderReplaceRequest BuildReplaceRequest ()
		{
			return new ContentControlFolderBatchReplaceService.FolderReplaceRequest {
				TemplateDirectory = (_folderTextBox.Text ?? string.Empty).Trim (),
				OldTag = _oldTagTextBox.Text,
				NewTag = _newTagTextBox.Text,
				OldTitle = _oldTitleTextBox.Text,
				NewTitle = _newTitleTextBox.Text,
				OldDisplayText = _oldDisplayTextBox.Text,
				NewDisplayText = _newDisplayTextBox.Text,
				UsePartialMatch = _partialMatchCheckBox.Checked,
				CreateBackups = _backupCheckBox.Checked
			};
		}

		private static Label CreateLabel (string text, int left, int top)
		{
			return new Label {
				Left = left,
				Top = top,
				Width = 92,
				Height = 22,
				Text = text
			};
		}

		private static TextBox CreateTextBox (int left, int top, int width)
		{
			return new TextBox {
				Left = left,
				Top = top,
				Width = width,
				Height = 24
			};
		}
	}
}
