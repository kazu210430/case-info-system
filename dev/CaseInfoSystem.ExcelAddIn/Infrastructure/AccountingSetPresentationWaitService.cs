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

			internal WaitSession (Logger logger)
			{
				_logger = logger ?? throw new ArgumentNullException ("logger");
			}

			internal void Show (Stopwatch commandStopwatch)
			{
				try {
					_waitForm = new AccountingSetPresentationWaitForm ();
					_waitForm.Show ();
					_waitForm.Activate ();
					_waitForm.BringToFront ();
					_waitForm.Update ();
					_waitForm.Refresh ();
					Application.DoEvents ();
					_logger.Info ("Accounting set presentation wait UI shown. elapsedMs=" + GetElapsedMilliseconds (commandStopwatch));
				} catch (Exception exception) {
					Close ();
					_logger.Warn ("Accounting set presentation wait UI failed to show. message=" + exception.Message);
				}
			}

			internal void UpdateStage (string title, string detail = null)
			{
				if (_isClosed || _waitForm == null || _waitForm.IsDisposed) {
					return;
				}

				try {
					_waitForm.SetStage (title, string.IsNullOrWhiteSpace (detail) ? DefaultStageDetail : detail);
					_waitForm.Update ();
					_waitForm.Refresh ();
					Application.DoEvents ();
				} catch (Exception exception) {
					_logger.Warn ("Accounting set presentation wait UI stage update failed. title=" + (title ?? string.Empty) + ", message=" + exception.Message);
				}
			}

			internal void Close ()
			{
				if (_isClosed) {
					return;
				}

				_isClosed = true;
				if (_waitForm == null) {
					return;
				}

				try {
					if (!_waitForm.IsDisposed) {
						_waitForm.Close ();
					}
				} catch (Exception exception) {
					_logger.Warn ("Accounting set presentation wait UI failed to close cleanly. message=" + exception.Message);
				} finally {
					_waitForm.Dispose ();
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
		}
	}
}
