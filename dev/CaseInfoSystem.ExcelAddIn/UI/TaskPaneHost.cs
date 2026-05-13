using System;
using System.Globalization;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	// Holds the concrete CustomTaskPane lifetime for one window-bound host.
	// Dispose() is the concrete lifetime release boundary, while host-map ordering and metadata timing stay outside this type.
	internal sealed class TaskPaneHost : IDisposable
	{
		private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

		private readonly ThisAddIn _addIn;

		private readonly UserControl _control;

		private readonly ITaskPaneView _view;

		private readonly Window _window;

		private readonly string _windowKey;

		private readonly Logger _logger;

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

		internal TaskPaneHost (ThisAddIn addIn, Window window, UserControl control, ITaskPaneView view, string windowKey, Logger logger = null)
		{
			_addIn = addIn ?? throw new ArgumentNullException ("addIn");
			_window = window ?? throw new ArgumentNullException ("window");
			_control = control ?? throw new ArgumentNullException ("control");
			_view = view ?? throw new ArgumentNullException ("view");
			_windowKey = windowKey ?? throw new ArgumentNullException ("windowKey");
			_logger = logger;
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
			LogInfo ("dispose-start");
			try {
				LogInfo ("hide-start");
				Hide ();
				LogInfo ("hide-complete");
			} catch (Exception exception) {
				LogFailure ("hide-failure", exception);
			}
			if (_pane == null) {
				LogInfo ("dispose-complete paneAlreadyNull=True");
				return;
			}
			try {
				LogInfo ("remove-taskpane-start");
				_addIn.RemoveTaskPane (_pane);
				LogInfo ("remove-taskpane-complete");
			} catch (Exception exception) {
				LogFailure ("remove-taskpane-failure", exception);
			} finally {
				_pane = null;
			}
			LogInfo ("dispose-complete paneAlreadyNull=False");
		}

		private void LogInfo (string action)
		{
			_logger?.Info (
				KernelFlickerTracePrefix
				+ " source=TaskPaneHost action="
				+ (action ?? string.Empty)
				+ " host="
				+ FormatSafeDescriptor ());
		}

		private void LogFailure (string action, Exception exception)
		{
			_logger?.Error (
				KernelFlickerTracePrefix
				+ " source=TaskPaneHost action="
				+ (action ?? string.Empty)
				+ " host="
				+ FormatSafeDescriptor ()
				+ ", exceptionType="
				+ exception.GetType ().Name
				+ ", hResult=0x"
				+ exception.HResult.ToString ("X8", CultureInfo.InvariantCulture)
				+ ", message="
				+ exception.Message,
				exception);
		}

		private string FormatSafeDescriptor ()
		{
			return "paneRole="
				+ GetSafePaneRoleName ()
				+ ", windowKey="
				+ (_windowKey ?? string.Empty)
				+ ", workbookFullName="
				+ (WorkbookFullName ?? string.Empty)
				+ ", controlType="
				+ (_control?.GetType ().Name ?? string.Empty);
		}

		private string GetSafePaneRoleName ()
		{
			if (_control is KernelNavigationControl) {
				return "Kernel";
			}
			if (_control is DocumentButtonsControl) {
				return "Case";
			}
			if (_control is AccountingNavigationControl) {
				return "Accounting";
			}
			return "Unknown";
		}
	}
}
