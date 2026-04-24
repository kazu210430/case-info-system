using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    internal sealed class DocumentPresentationWaitForm : Form
    {
        internal DocumentPresentationWaitForm()
        {
            Text = "案件情報System";
            Font = new Font("Yu Gothic UI", 10f, FontStyle.Regular, GraphicsUnit.Point, 128);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            TopMost = true;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            DoubleBuffered = true;
            ClientSize = new Size(392, 116);

            Label titleLabel = new Label
            {
                AutoSize = false,
                Left = 20,
                Top = 22,
                Width = 352,
                Height = 26,
                Text = "文書を開く準備をしています",
                TextAlign = ContentAlignment.MiddleLeft
            };

            Label detailLabel = new Label
            {
                AutoSize = false,
                Left = 20,
                Top = 54,
                Width = 352,
                Height = 18,
                Text = "Word の起動や保存が完了するまで、そのままお待ちください。",
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };

            Controls.Add(titleLabel);
            Controls.Add(detailLabel);
        }
    }
}
