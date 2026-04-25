using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class CreatedCasePresentationWaitForm : Form
	{
		private readonly Label _titleLabel;

		private readonly Label _detailLabel;

		internal CreatedCasePresentationWaitForm ()
		{
			Text = "案件情報System";
			Font = new Font ("Yu Gothic UI", 10f, FontStyle.Regular, GraphicsUnit.Point, 128);
			FormBorderStyle = FormBorderStyle.FixedDialog;
			StartPosition = FormStartPosition.CenterScreen;
			ShowInTaskbar = false;
			TopMost = true;
			ControlBox = false;
			MaximizeBox = false;
			MinimizeBox = false;
			DoubleBuffered = true;
			ClientSize = new Size (360, 116);
			_titleLabel = new Label {
				AutoSize = false,
				Left = 20,
				Top = 22,
				Width = 320,
				Height = 26,
				Text = "案件情報.xlsxを作成しています",
				TextAlign = ContentAlignment.MiddleLeft
			};
			_detailLabel = new Label {
				AutoSize = false,
				Left = 20,
				Top = 54,
				Width = 320,
				Height = 18,
				Text = "画面が切り替わるまでそのままでお待ちください。",
				ForeColor = Color.DimGray,
				TextAlign = ContentAlignment.MiddleLeft
			};
			Controls.Add (_titleLabel);
			Controls.Add (_detailLabel);
		}

		internal void SetStage (string title, string detail)
		{
			if (IsDisposed) {
				return;
			}
			if (InvokeRequired) {
				Invoke ((MethodInvoker)delegate {
					SetStageCore (title, detail);
				});
				return;
			}
			SetStageCore (title, detail);
		}

		private void SetStageCore (string title, string detail)
		{
			if (IsDisposed) {
				return;
			}
			_titleLabel.Text = string.IsNullOrWhiteSpace (title) ? _titleLabel.Text : title;
			_detailLabel.Text = string.IsNullOrWhiteSpace (detail) ? _detailLabel.Text : detail;
		}
	}
}
