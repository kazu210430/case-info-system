using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class TaskPaneHost : IDisposable
	{
		private readonly ThisAddIn _addIn;

		private readonly UserControl _control;

		private readonly ITaskPaneView _view;

		private readonly Window _window;

		private readonly string _windowKey;

		private CustomTaskPane _pane;

		internal UserControl Control => _control;

		internal string WindowKey => _windowKey;

		internal Window Window => _window;

		internal string WorkbookFullName { get; set; }

		internal string LastRenderSignature { get; set; }

		internal TaskPaneHost (ThisAddIn addIn, Window window, UserControl control, ITaskPaneView view, string windowKey)
		{
			_addIn = addIn ?? throw new ArgumentNullException ("addIn");
			_window = window ?? throw new ArgumentNullException ("window");
			_control = control ?? throw new ArgumentNullException ("control");
			_view = view ?? throw new ArgumentNullException ("view");
			_windowKey = windowKey ?? throw new ArgumentNullException ("windowKey");
			_pane = _addIn.CreateTaskPane (window, _control);
		}

		internal void Show ()
		{
			if (_pane != null) {
				_pane.Width = _view.PreferredWidth;
				_pane.Visible = true;
			}
		}

		internal void Hide ()
		{
			if (_pane != null) {
				_pane.Visible = false;
			}
		}

		public void Dispose ()
		{
			try {
				Hide ();
			} catch {
			}
			if (_pane == null) {
				return;
			}
			try {
				_addIn.RemoveTaskPane (_pane);
			} catch {
			} finally {
				_pane = null;
			}
		}
	}
}
