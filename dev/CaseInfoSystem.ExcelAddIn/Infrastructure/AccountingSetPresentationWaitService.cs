using System;
using System.Diagnostics;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class AccountingSetPresentationWaitService
	{
		internal const string CreatingStageTitle = "会計書類を作成しています";

		internal const string OpeningWorkbookStageTitle = "会計書類を開いています";

		internal const string ApplyingInitialDataStageTitle = "初期データを設定しています";

		internal const string ShowingInputScreenStageTitle = "入力画面を表示しています";

		internal const string DefaultStageDetail = "画面が表示されるまで、そのままでお待ちください。";

		private readonly Logger _logger;

		internal AccountingSetPresentationWaitService (Logger logger)
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

			private AccountingSetPresentationWaitForm _waitForm;

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
					_waitForm = new AccountingSetPresentationWaitForm ();
					_waitForm.Show ();
					_waitForm.Activate ();
					_waitForm.BringToFront ();
					RefreshWaitForm (_waitForm);
					_logger.Info ("Accounting set presentation wait UI shown. elapsedMs=" + GetElapsedMilliseconds (commandStopwatch));
				} catch (Exception exception) {
					Close ();
					_logger.Warn ("Accounting set presentation wait UI failed to show. message=" + exception.Message);
				} finally {
					_isShowing = false;
				}
			}

			internal void UpdateStage (string title, string detail = null)
			{
				if (_isClosed || _isClosing || _isUpdating) {
					return;
				}

				try {
					AccountingSetPresentationWaitForm waitForm = GetActiveWaitForm ();
					if (waitForm == null) {
						return;
					}

					_isUpdating = true;
					waitForm.SetStage (title, string.IsNullOrWhiteSpace (detail) ? DefaultStageDetail : detail);
					RefreshWaitForm (waitForm);
				} catch (Exception exception) {
					_logger.Warn ("Accounting set presentation wait UI stage update failed. title=" + (title ?? string.Empty) + ", message=" + exception.Message);
				} finally {
					_isUpdating = false;
				}
			}

			internal void Close ()
			{
				if (_isClosed || _isClosing) {
					return;
				}

				_isClosed = true;
				_isClosing = true;
				AccountingSetPresentationWaitForm waitForm = _waitForm;
				if (waitForm == null) {
					return;
				}

				try {
					if (!waitForm.IsDisposed) {
						if (waitForm.Visible) {
							waitForm.Close ();
						} else {
							waitForm.Dispose ();
						}
					}
				} catch (Exception exception) {
					_logger.Warn ("Accounting set presentation wait UI failed to close cleanly. message=" + exception.Message);
				} finally {
					if (!waitForm.IsDisposed) {
						waitForm.Dispose ();
					}
					_waitForm = null;
				}
			}

			public void Dispose ()
			{
				Close ();
			}

			private static long GetElapsedMilliseconds (Stopwatch stopwatch)
			{
				return (stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds;
			}

			private static void RefreshWaitForm (AccountingSetPresentationWaitForm waitForm)
			{
				if (waitForm == null || waitForm.IsDisposed) {
					return;
				}

				if (waitForm.IsHandleCreated) {
					waitForm.Invalidate ();
					waitForm.Update ();
					waitForm.Refresh ();
					try {
						waitForm.BeginInvoke ((MethodInvoker)(() =>
						{
							if (waitForm.IsDisposed) {
								return;
							}

							waitForm.Invalidate ();
							waitForm.Update ();
							waitForm.Refresh ();
						}));
					} catch (InvalidOperationException) {
					}
				}
			}

			private AccountingSetPresentationWaitForm GetActiveWaitForm ()
			{
				return (_waitForm == null || _waitForm.IsDisposed) ? null : _waitForm;
			}
		}
	}
}
