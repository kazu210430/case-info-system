using System;
using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class CreatedCaseOpenPromptForm : Form
	{
		private const string PromptCaption = "案件情報System";

		private const string PromptMessage = "案件情報を開きますか？";

		private readonly Action _shownCallback;

		private CreatedCaseOpenPromptForm (Action shownCallback)
		{
			_shownCallback = shownCallback;
			Text = "案件情報System";
			Font = new Font ("Yu Gothic UI", 10f, FontStyle.Regular, GraphicsUnit.Point, 128);
			base.FormBorderStyle = FormBorderStyle.FixedDialog;
			base.StartPosition = FormStartPosition.CenterScreen;
			base.MaximizeBox = false;
			base.MinimizeBox = false;
			base.ShowInTaskbar = false;
			base.TopMost = true;
			DoubleBuffered = true;
			base.ClientSize = new Size (360, 130);
			Label value = new Label {
				AutoSize = false,
				Left = 20,
				Top = 20,
				Width = 320,
				Height = 32,
				Text = "案件情報を開きますか？",
				TextAlign = ContentAlignment.MiddleLeft
			};
			Button button = new Button {
				Left = 164,
				Top = 76,
				Width = 84,
				Text = "はい",
				DialogResult = DialogResult.Yes
			};
			Button button2 = new Button {
				Left = 256,
				Top = 76,
				Width = 84,
				Text = "いいえ",
				DialogResult = DialogResult.No
			};
			base.AcceptButton = button;
			base.CancelButton = button2;
			base.Controls.Add (value);
			base.Controls.Add (button);
			base.Controls.Add (button2);
		}

		protected override void OnShown (EventArgs e)
		{
			base.OnShown (e);
			_shownCallback?.Invoke ();
		}

		internal static CreatedCaseOpenDecision ShowPrompt (IWin32Window owner, Action shownCallback)
		{
			using (CreatedCaseOpenPromptForm createdCaseOpenPromptForm = new CreatedCaseOpenPromptForm (shownCallback)) {
				DialogResult dialogResult = ((owner == null) ? createdCaseOpenPromptForm.ShowDialog () : createdCaseOpenPromptForm.ShowDialog (owner));
				return (dialogResult != DialogResult.Yes) ? CreatedCaseOpenDecision.Skip : CreatedCaseOpenDecision.Open;
			}
		}
	}
}
