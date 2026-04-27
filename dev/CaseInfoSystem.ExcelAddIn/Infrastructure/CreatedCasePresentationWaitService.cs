using System;
using System.Diagnostics;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class CreatedCasePresentationWaitService
	{
		internal const string CreatingStageTitle = "案件情報.xlsxを作成しています";

		internal const string PreparingOpenStageTitle = "案件情報.xlsxを開く準備をしています";

		internal const string ShowingScreenStageTitle = "案件情報.xlsxの画面を表示しています";

		internal const string BatchOpeningFolderStageTitle = "保存先フォルダを開いています";

		internal const string BatchReturningHomeStageTitle = "HOME画面に戻ります。作成を続けてください。";

		internal const string DefaultStageDetail = "画面が切り替わるまでそのままでお待ちください。";

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

			private bool _isShowing;

			private bool _isUpdating;

			private bool _isClosing;

			internal WaitSession (Logger logger)
			{
				_logger = logger ?? throw new ArgumentNullException ("logger");
			}

			internal void Show (Stopwatch commandStopwatch)
			{
				if (_isClosed || _isShowing || _isClosing) {
					return;
				}
				try {
					_isShowing = true;
					PrepareOwnerForWaitDisplay ();
					_waitForm = new CreatedCasePresentationWaitForm ();
					_waitForm.Show ();
					_waitForm.Activate ();
					_waitForm.BringToFront ();
					RefreshWaitForm (_waitForm);
					_logger.Info ("Created CASE presentation wait UI shown. elapsedMs=" + GetElapsedMilliseconds (commandStopwatch));
				} catch (Exception exception) {
					CloseCore (restoreOwner: true);
					_logger.Warn ("Created CASE presentation wait UI failed to show. message=" + exception.Message);
				} finally {
					_isShowing = false;
				}
			}

			internal void CloseForSuccessfulPresentation ()
			{
				CloseCore (restoreOwner: false);
			}

			internal void UpdateStage (string title, string detail = null)
			{
				if (_isClosed || _isClosing || _isUpdating) {
					return;
				}
				try {
					CreatedCasePresentationWaitForm waitForm = GetActiveWaitForm ();
					if (waitForm == null) {
						return;
					}
					_isUpdating = true;
					waitForm.SetStage (title, string.IsNullOrWhiteSpace (detail) ? DefaultStageDetail : detail);
					RefreshWaitForm (waitForm);
				} catch (Exception exception) {
					_logger.Warn ("Created CASE presentation wait UI stage update failed. title=" + (title ?? string.Empty) + ", message=" + exception.Message);
				} finally {
					_isUpdating = false;
				}
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
				if (_isClosed || _isClosing) {
					return;
				}
				_isClosed = true;
				_isClosing = true;
				CreatedCasePresentationWaitForm waitForm = _waitForm;
				if (waitForm != null) {
					try {
						if (!waitForm.IsDisposed) {
							if (waitForm.Visible) {
								waitForm.Close ();
							} else {
								waitForm.Dispose ();
							}
						}
					} catch (Exception exception) {
						_logger.Warn ("Created CASE presentation wait UI failed to close cleanly. message=" + exception.Message);
					} finally {
						if (!waitForm.IsDisposed) {
							waitForm.Dispose ();
						}
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

			private static void RefreshWaitForm (CreatedCasePresentationWaitForm waitForm)
			{
				if (waitForm == null || waitForm.IsDisposed) {
					return;
				}
				if (waitForm.IsHandleCreated) {
					waitForm.Invalidate ();
					waitForm.Update ();
					waitForm.Refresh ();
				}
			}

			private CreatedCasePresentationWaitForm GetActiveWaitForm ()
			{
				return (_waitForm == null || _waitForm.IsDisposed) ? null : _waitForm;
			}
		}
	}
}
