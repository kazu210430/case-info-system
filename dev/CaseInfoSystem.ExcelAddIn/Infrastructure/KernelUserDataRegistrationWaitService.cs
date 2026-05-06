using System;
using System.Diagnostics;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class KernelUserDataRegistrationWaitService : IKernelUserDataRegistrationWaitService
    {
        internal const string StageTitle = "ユーザー情報を登録しています...";
        internal const string StageDetail = "転記が終わるまで、そのままでお待ちください。";

        private readonly Logger _logger;

        internal KernelUserDataRegistrationWaitService(Logger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public IDisposable ShowWaiting(Stopwatch commandStopwatch)
        {
            WaitSession session = new WaitSession(_logger);
            session.Show(commandStopwatch);
            return session;
        }

        private sealed class WaitSession : IDisposable
        {
            private readonly Logger _logger;
            private KernelUserDataRegistrationWaitForm _waitForm;
            private bool _isClosed;
            private bool _isShowing;
            private bool _isClosing;

            internal WaitSession(Logger logger)
            {
                _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            }

            internal void Show(Stopwatch commandStopwatch)
            {
                if (_isClosed || _isShowing || _isClosing)
                {
                    return;
                }

                try
                {
                    _isShowing = true;
                    _waitForm = new KernelUserDataRegistrationWaitForm();
                    _waitForm.Show();
                    _waitForm.Activate();
                    _waitForm.BringToFront();
                    RefreshWaitForm(_waitForm);
                    _logger.Info("Kernel user data registration wait UI shown. elapsedMs=" + GetElapsedMilliseconds(commandStopwatch));
                }
                catch (Exception exception)
                {
                    Dispose();
                    _logger.Warn("Kernel user data registration wait UI failed to show. message=" + exception.Message);
                }
                finally
                {
                    _isShowing = false;
                }
            }

            public void Dispose()
            {
                if (_isClosed || _isClosing)
                {
                    return;
                }

                _isClosed = true;
                _isClosing = true;
                KernelUserDataRegistrationWaitForm waitForm = _waitForm;
                if (waitForm == null)
                {
                    return;
                }

                try
                {
                    if (!waitForm.IsDisposed)
                    {
                        if (waitForm.Visible)
                        {
                            waitForm.Close();
                        }
                        else
                        {
                            waitForm.Dispose();
                        }
                    }
                }
                catch (Exception exception)
                {
                    _logger.Warn("Kernel user data registration wait UI failed to close cleanly. message=" + exception.Message);
                }
                finally
                {
                    if (!waitForm.IsDisposed)
                    {
                        waitForm.Dispose();
                    }

                    _waitForm = null;
                }
            }

            private static long GetElapsedMilliseconds(Stopwatch stopwatch)
            {
                return stopwatch == null ? 0L : stopwatch.ElapsedMilliseconds;
            }

            private static void RefreshWaitForm(KernelUserDataRegistrationWaitForm waitForm)
            {
                if (waitForm == null || waitForm.IsDisposed)
                {
                    return;
                }

                if (waitForm.IsHandleCreated)
                {
                    waitForm.Invalidate();
                    waitForm.Update();
                    waitForm.Refresh();
                    try
                    {
                        waitForm.BeginInvoke((MethodInvoker)delegate
                        {
                            if (waitForm.IsDisposed)
                            {
                                return;
                            }

                            waitForm.Invalidate();
                            waitForm.Update();
                            waitForm.Refresh();
                        });
                    }
                    catch (InvalidOperationException)
                    {
                    }
                }
            }
        }
    }
}
