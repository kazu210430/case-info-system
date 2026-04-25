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

            internal void UpdateStage(string title, string detail = null)
            {
                if (_isClosed || _waitForm == null || _waitForm.IsDisposed)
                {
                    return;
                }

                try
                {
                    _waitForm.SetStage(title, string.IsNullOrWhiteSpace(detail) ? DefaultStageDetail : detail);
                    _waitForm.Update();
                    _waitForm.Refresh();
                    Application.DoEvents();
                }
                catch (Exception exception)
                {
                    _logger.Warn("Document presentation wait UI stage update failed. title=" + (title ?? string.Empty) + ", message=" + exception.Message);
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
