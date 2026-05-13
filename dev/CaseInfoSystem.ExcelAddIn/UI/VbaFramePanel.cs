using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    /// <summary>
    internal sealed class VbaFramePanel : Panel
    {
        private string _caption = string.Empty;

        internal VbaFramePanel()
        {
            BackColor = Color.FromArgb(229, 245, 255);
            Font = new Font("Yu Gothic UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 128);
            TabStop = false;
            ResizeRedraw = true;
        }

        /// <summary>
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        internal string Caption
        {
            get { return _caption; }
            set
            {
                _caption = value ?? string.Empty;
                Invalidate();
            }
        }

        /// <summary>
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            TextRenderer.DrawText(
                e.Graphics,
                _caption,
                Font,
                new Point(8, 0),
                ForeColor,
                BackColor);

            Size captionSize = TextRenderer.MeasureText(e.Graphics, _caption, Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
            int captionWidth = Math.Max(0, captionSize.Width);
            int borderY = Font.Height / 2;

            using (var pen = new Pen(SystemColors.ControlDark))
            {
                e.Graphics.DrawLine(pen, 0, borderY, 6, borderY);
                e.Graphics.DrawLine(pen, 8 + captionWidth, borderY, Width - 1, borderY);
                e.Graphics.DrawLine(pen, 0, borderY, 0, Height - 1);
                e.Graphics.DrawLine(pen, Width - 1, borderY, Width - 1, Height - 1);
                e.Graphics.DrawLine(pen, 0, Height - 1, Width - 1, Height - 1);
            }
        }
    }
}
