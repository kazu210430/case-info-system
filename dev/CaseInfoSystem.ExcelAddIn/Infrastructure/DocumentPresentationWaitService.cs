using System;
using System.Diagnostics;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class DocumentPresentationWaitService
    {
        internal const string CreatingStageTitle = "文書を作成しています";
        internal const string LaunchingWordStageTitle = "Wordを起動しています";
        internal const string LoadingTemplateStageTitle = "テンプレートを読み込んでいます";
        internal const string ApplyingMergeDataStageTitle = "文書へ差し込んでいます";
        internal const string SavingDocumentStageTitle = "文書を保存しています";
        internal const string ShowingScreenStageTitle = "画面を表示しています";
        internal const string DefaultStageDetail = "Word の起動や処理が完了するまで、そのままでお待ちください。";

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
            private bool _isShowing;
            private bool _isUpdating;
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
                    _waitForm = new DocumentPresentationWaitForm();
                    _waitForm.Show();
                    _waitForm.Activate();
                    _waitForm.BringToFront();
                    RefreshWaitForm(_waitForm);
                    _logger.Info("Document presentation wait UI shown. elapsedMs=" + GetElapsedMilliseconds(commandStopwatch));
                }
                catch (Exception exception)
                {
                    Close();
                    _logger.Warn("Document presentation wait UI failed to show. message=" + exception.Message);
                }
                finally
                {
                    _isShowing = false;
                }
            }

            internal void UpdateStage(string title, string detail = null)
            {
                if (_isClosed || _isClosing || _isUpdating)
                {
                    return;
                }

                try
                {
                    DocumentPresentationWaitForm waitForm = GetActiveWaitForm();
                    if (waitForm == null)
                    {
                        return;
                    }

                    _isUpdating = true;
                    waitForm.SetStage(title, string.IsNullOrWhiteSpace(detail) ? DefaultStageDetail : detail);
                    RefreshWaitForm(waitForm);
                }
                catch (Exception exception)
                {
                    _logger.Warn("Document presentation wait UI stage update failed. title=" + (title ?? string.Empty) + ", message=" + exception.Message);
                }
                finally
                {
                    _isUpdating = false;
                }
            }

            internal void Close()
            {
                if (_isClosed || _isClosing)
                {
                    return;
                }

                _isClosed = true;
                _isClosing = true;
                DocumentPresentationWaitForm waitForm = _waitForm;
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
                    _logger.Warn("Document presentation wait UI failed to close cleanly. message=" + exception.Message);
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

            public void Dispose()
            {
                Close();
            }

            private static long GetElapsedMilliseconds(Stopwatch stopwatch)
            {
                return stopwatch == null ? 0L : stopwatch.ElapsedMilliseconds;
            }

            private static void RefreshWaitForm(DocumentPresentationWaitForm waitForm)
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
                }
            }

            private DocumentPresentationWaitForm GetActiveWaitForm()
            {
                return _waitForm == null || _waitForm.IsDisposed ? null : _waitForm;
            }
        }
    }
}
