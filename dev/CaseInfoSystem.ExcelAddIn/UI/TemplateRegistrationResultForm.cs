using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class TemplateRegistrationResultForm : Form
	{
		private readonly TextBox _messageTextBox;

		private readonly Button _okButton;

		internal TemplateRegistrationResultForm (string title, string message)
		{
			Text = string.IsNullOrWhiteSpace (title) ? "案件情報System" : title;
			StartPosition = FormStartPosition.CenterScreen;
			FormBorderStyle = FormBorderStyle.FixedDialog;
			MaximizeBox = false;
			MinimizeBox = false;
			ShowInTaskbar = false;
			Font = new Font ("Yu Gothic UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 128);
			ClientSize = new Size (760, 560);

			_messageTextBox = new TextBox {
				Location = new Point (20, 20),
				Size = new Size (720, 470),
				Multiline = true,
				ReadOnly = true,
				ScrollBars = ScrollBars.Vertical,
				WordWrap = true,
				BorderStyle = BorderStyle.FixedSingle,
				Text = message ?? string.Empty
			};

			_okButton = new Button {
				Text = "OK",
				DialogResult = DialogResult.OK,
				Size = new Size (96, 32),
				Location = new Point (644, 508)
			};

			Controls.Add (_messageTextBox);
			Controls.Add (_okButton);

			AcceptButton = _okButton;
			CancelButton = _okButton;
		}

		internal static DialogResult ShowNotice (string title, string message)
		{
			using (TemplateRegistrationResultForm templateRegistrationResultForm = new TemplateRegistrationResultForm (title, message)) {
				return templateRegistrationResultForm.ShowDialog ();
			}
		}
	}
}
