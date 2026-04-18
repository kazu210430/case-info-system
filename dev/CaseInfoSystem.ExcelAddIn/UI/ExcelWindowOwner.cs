using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class ExcelWindowOwner : NativeWindow, IWin32Window, IDisposable
	{
		private bool _assigned;

		internal static ExcelWindowOwner From (Window window)
		{
			if (window == null) {
				return null;
			}
			try {
				int hwnd = window.Hwnd;
				if (hwnd == 0) {
					return null;
				}
				ExcelWindowOwner excelWindowOwner = new ExcelWindowOwner ();
				excelWindowOwner.AssignHandle (new IntPtr (hwnd));
				excelWindowOwner._assigned = true;
				return excelWindowOwner;
			} catch {
				return null;
			}
		}

		public void Dispose ()
		{
			if (_assigned) {
				ReleaseHandle ();
				_assigned = false;
			}
		}
	}
}
