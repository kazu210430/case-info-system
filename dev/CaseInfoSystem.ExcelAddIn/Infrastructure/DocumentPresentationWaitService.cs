using System;
using System.Diagnostics;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class DocumentPresentationWaitService
    {
        private readonly Logger _logger;

        internal DocumentPresentationWaitService(Logger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal WaitSession ShowWaiting(Stopwatch commandStopwatch)
        {
            WaitSession waitSession = new WaitSession(_logger);
            waitSession.Show(commandStopwatch);
            return waitSession;
        }

        internal sealed class WaitSession : IDisposable
        {
            private readonly Logger _logger;
            private DocumentPresentationWaitForm _waitForm;
            private bool _isClosed;

            internal WaitSession(Logger logger)
            {
                _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            }

            internal void Show(Stopwatch commandStopwatch)
            {
                try
                {
                    _waitForm = new DocumentPresentationWaitForm();
                    _waitForm.Show();
                    _waitForm.Activate();
                    _waitForm.BringToFront();
                    _waitForm.Update();
                    _waitForm.Refresh();
                    Application.DoEvents();
                    _logger.Info("Document presentation wait UI shown. elapsedMs=" + GetElapsedMilliseconds(commandStopwatch));
                }
                catch (Exception exception)
                {
                    Close();
                    _logger.Warn("Document presentation wait UI failed to show. message=" + exception.Message);
                }
            }

            internal void Close()
            {
                if (_isClosed)
                {
                    return;
                }

                _isClosed = true;
                if (_waitForm == null)
                {
                    return;
                }

                try
                {
                    if (!_waitForm.IsDisposed)
                    {
                        _waitForm.Close();
                    }
                }
                catch (Exception exception)
                {
                    _logger.Warn("Document presentation wait UI failed to close cleanly. message=" + exception.Message);
                }
                finally
                {
                    _waitForm.Dispose();
                    _waitForm = null;
                }
            }

            public void Dispose()
            {
                Close();
            }

            private static long GetElapsedMilliseconds(Stopwatch stopwatch)
            {
                return stopwatch == null ? 0L : stopwatch.ElapsedMilliseconds;
            }
        }
    }
}
