using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    /// <summary>
    internal static class ButtonCursorHelper
    {
        /// <summary>
        internal static void ApplyHandCursor(Control root)
        {
            if (root == null)
            {
                return;
            }

            if (root is ButtonBase)
            {
                root.Cursor = Cursors.Hand;
            }

            foreach (Control child in root.Controls)
            {
                ApplyHandCursor(child);
            }
        }
    }
}
