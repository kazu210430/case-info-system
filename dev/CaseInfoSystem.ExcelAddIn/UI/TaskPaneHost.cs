using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	// Holds the concrete CustomTaskPane lifetime for one window-bound host.
	// Dispose() is the concrete lifetime release boundary, while host-map ordering and metadata timing stay outside this type.
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

		internal bool IsVisible {
			get {
				try {
					return _pane != null && _pane.Visible;
				} catch {
					return false;
				}
			}
		}

		internal TaskPaneHost (ThisAddIn addIn, Window window, UserControl control, ITaskPaneView view, string windowKey)
		{
			_addIn = addIn ?? throw new ArgumentNullException ("addIn");
			_window = window ?? throw new ArgumentNullException ("window");
			_control = control ?? throw new ArgumentNullException ("control");
			_view = view ?? throw new ArgumentNullException ("view");
			_windowKey = windowKey ?? throw new ArgumentNullException ("windowKey");
			// This is the point where host lifetime crosses the VSTO adapter boundary into ThisAddIn.
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
			// Concrete lifetime release boundary for a host that is already being removed by outer orchestration.
			// Current-state fixed point is Hide() -> RemoveTaskPane(...) -> _pane = null.
			// Event unbinding stays implicit via disposal and is not clarified here.
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
