using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class CreatedCasePresentationWaitForm : Form
	{
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
			Label label = new Label {
				AutoSize = false,
				Left = 20,
				Top = 22,
				Width = 320,
				Height = 26,
				Text = "案件情報.xlsxを起動しています",
				TextAlign = ContentAlignment.MiddleLeft
			};
			Label label2 = new Label {
				AutoSize = false,
				Left = 20,
				Top = 54,
				Width = 320,
				Height = 18,
				Text = "画面が切り替わるまでそのままでお待ちください。",
				ForeColor = Color.DimGray,
				TextAlign = ContentAlignment.MiddleLeft
			};
			Controls.Add (label);
			Controls.Add (label2);
		}
	}
}
