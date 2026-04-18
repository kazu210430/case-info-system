using System;
using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    /// <summary>
    internal sealed class CompletionNoticeForm : Form
    {
        private readonly Label _messageLabel;
        private readonly Button _okButton;

        /// <summary>
        internal CompletionNoticeForm(string title, string message)
        {
            Text = string.IsNullOrWhiteSpace(title) ? "\u6848\u4EF6\u60C5\u5831System" : title;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Yu Gothic UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 128);
            ClientSize = new Size(420, 150);

            _messageLabel = new Label
            {
                AutoSize = false,
                Location = new Point(24, 24),
                Size = new Size(372, 52),
                Text = message ?? string.Empty,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Size = new Size(96, 32),
                Location = new Point(300, 94)
            };

            Controls.Add(_messageLabel);
            Controls.Add(_okButton);

            AcceptButton = _okButton;
            CancelButton = _okButton;
        }

        /// <summary>
        internal static DialogResult ShowNotice(IWin32Window owner, string title, string message)
        {
            using (var form = new CompletionNoticeForm(title, message))
            {
                return owner == null ? form.ShowDialog() : form.ShowDialog(owner);
            }
        }
    }
}
