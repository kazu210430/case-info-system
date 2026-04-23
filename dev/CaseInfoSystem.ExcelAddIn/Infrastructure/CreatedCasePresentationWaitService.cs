using System;
using System.Diagnostics;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class CreatedCasePresentationWaitService
	{
		private readonly Logger _logger;

		internal CreatedCasePresentationWaitService (Logger logger)
		{
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal WaitSession ShowWaiting (Stopwatch commandStopwatch)
		{
			WaitSession waitSession = new WaitSession (_logger);
			waitSession.Show (commandStopwatch);
			return waitSession;
		}

		internal sealed class WaitSession : IDisposable
		{
			private readonly Logger _logger;

			private Form _temporarilyMinimizedOwner;

			private FormWindowState _temporarilyMinimizedOwnerState = FormWindowState.Normal;

			private CreatedCasePresentationWaitForm _waitForm;

			private bool _isClosed;

			internal WaitSession (Logger logger)
			{
				_logger = logger ?? throw new ArgumentNullException ("logger");
			}

			internal void Show (Stopwatch commandStopwatch)
			{
				try {
					PrepareOwnerForWaitDisplay ();
					_waitForm = new CreatedCasePresentationWaitForm ();
					_waitForm.Show ();
					_waitForm.Activate ();
					_waitForm.BringToFront ();
					_waitForm.Update ();
					_waitForm.Refresh ();
					System.Windows.Forms.Application.DoEvents ();
					_logger.Info ("Created CASE presentation wait UI shown. elapsedMs=" + GetElapsedMilliseconds (commandStopwatch));
				} catch (Exception exception) {
					CloseCore (restoreOwner: true);
					_logger.Warn ("Created CASE presentation wait UI failed to show. message=" + exception.Message);
				}
			}

			internal void CloseForSuccessfulPresentation ()
			{
				CloseCore (restoreOwner: false);
			}

			internal void CloseAndRestoreOwner ()
			{
				CloseCore (restoreOwner: true);
			}

			public void Dispose ()
			{
				CloseCore (restoreOwner: true);
			}

			private void PrepareOwnerForWaitDisplay ()
			{
				ClearTrackedOwner ();
				if (ResolveWaitOwner () is Form form && !form.IsDisposed && form.Visible) {
					_temporarilyMinimizedOwner = form;
					_temporarilyMinimizedOwnerState = form.WindowState;
					form.WindowState = FormWindowState.Minimized;
				}
			}

			private void CloseCore (bool restoreOwner)
			{
				if (_isClosed) {
					return;
				}
				_isClosed = true;
				if (_waitForm != null) {
					try {
						if (!_waitForm.IsDisposed) {
							_waitForm.Close ();
						}
					} catch (Exception exception) {
						_logger.Warn ("Created CASE presentation wait UI failed to close cleanly. message=" + exception.Message);
					} finally {
						_waitForm.Dispose ();
						_waitForm = null;
					}
				}
				if (restoreOwner) {
					RestoreOwnerIfNeeded ();
				} else {
					ClearTrackedOwner ();
				}
			}

			private void RestoreOwnerIfNeeded ()
			{
				if (_temporarilyMinimizedOwner == null || _temporarilyMinimizedOwner.IsDisposed) {
					ClearTrackedOwner ();
					return;
				}
				try {
					_temporarilyMinimizedOwner.WindowState = _temporarilyMinimizedOwnerState;
					_temporarilyMinimizedOwner.Activate ();
				} catch (Exception exception) {
					_logger.Warn ("Created CASE presentation wait UI failed to restore owner. message=" + exception.Message);
				} finally {
					ClearTrackedOwner ();
				}
			}

			private void ClearTrackedOwner ()
			{
				_temporarilyMinimizedOwner = null;
				_temporarilyMinimizedOwnerState = FormWindowState.Normal;
			}

			private static IWin32Window ResolveWaitOwner ()
			{
				FormCollection openForms = System.Windows.Forms.Application.OpenForms;
				if (openForms == null || openForms.Count == 0) {
					return null;
				}
				for (int num = openForms.Count - 1; num >= 0; num--) {
					Form form = openForms [num];
					if (form != null && !form.IsDisposed && form.Visible) {
						return form;
					}
				}
				return null;
			}

			private static long GetElapsedMilliseconds (Stopwatch stopwatch)
			{
				return (stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds;
			}
		}
	}
}
