using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    internal sealed class KernelUserDataRegistrationWaitForm : Form
    {
        private readonly Label _titleLabel;
        private readonly Label _detailLabel;

        internal KernelUserDataRegistrationWaitForm()
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

            _titleLabel = new Label
            {
                AutoSize = false,
                Left = 20,
                Top = 22,
                Width = 352,
                Height = 26,
                Text = KernelUserDataRegistrationWaitService.StageTitle,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _detailLabel = new Label
            {
                AutoSize = false,
                Left = 20,
                Top = 54,
                Width = 352,
                Height = 18,
                Text = KernelUserDataRegistrationWaitService.StageDetail,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };

            Controls.Add(_titleLabel);
            Controls.Add(_detailLabel);
        }
    }
}
