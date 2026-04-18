using System;
using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    /// <summary>
    /// Class: lightweight text viewer form.
    /// Responsibility: show diagnostic or maintenance text in a readable modal dialog.
    /// </summary>
    internal sealed class TextDisplayForm : Form
    {
        private readonly TextBox _contentTextBox;
        private readonly Button _closeButton;
        private readonly Label _captionLabel;

        /// <summary>
        /// Method: initializes the form.
        /// Args: title - form title, caption - target label, content - body text.
        /// Returns: none.
        /// Side effects: creates and arranges WinForms controls.
        /// </summary>
        internal TextDisplayForm(string title, string caption, string content)
        {
            Text = string.IsNullOrWhiteSpace(title) ? "\u6848\u4EF6\u60C5\u5831System" : title;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(720, 420);
            Size = new Size(860, 560);
            Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 128);

            _captionLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                Padding = new Padding(12, 12, 12, 0),
                Text = string.IsNullOrWhiteSpace(caption) ? "\u5BFE\u8C61\u30D6\u30C3\u30AF" : caption
            };

            _contentTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 10F, FontStyle.Regular, GraphicsUnit.Point, 0),
                Text = content ?? string.Empty
            };

            _closeButton = new Button
            {
                Text = "\u9589\u3058\u308B",
                DialogResult = DialogResult.OK,
                Width = 96,
                Height = 32,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            };

            var buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                Padding = new Padding(12, 8, 12, 12)
            };

            buttonPanel.Controls.Add(_closeButton);
            _closeButton.Location = new Point(buttonPanel.Width - _closeButton.Width, 8);
            buttonPanel.Resize += ButtonPanel_Resize;

            Controls.Add(_contentTextBox);
            Controls.Add(buttonPanel);
            Controls.Add(_captionLabel);

            AcceptButton = _closeButton;
        }

        /// <summary>
        /// Method: keeps the close button right-aligned on panel resize.
        /// Args: sender - source panel, e - event args.
        /// Returns: none.
        /// Side effects: updates button position.
        /// </summary>
        private void ButtonPanel_Resize(object sender, EventArgs e)
        {
            Control panel = sender as Control;
            if (panel == null)
            {
                return;
            }

            _closeButton.Location = new Point(panel.ClientSize.Width - _closeButton.Width, 8);
        }
    }
}
