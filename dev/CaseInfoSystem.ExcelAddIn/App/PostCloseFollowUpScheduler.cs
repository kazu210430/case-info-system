using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class PostCloseFollowUpScheduler
    {
        private const int ExcelBusyHResult = unchecked((int)0x800AC472);
        private const int PostCloseRetryCount = 20;
        private const int PostCloseRetryIntervalMs = 500;

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;
        private readonly Queue<PostCloseFollowUpRequest> _pendingPostCloseQueue = new Queue<PostCloseFollowUpRequest>();
        private Control _dispatcher;
        private bool _postClosePosted;

        internal PostCloseFollowUpScheduler(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal void Schedule(string workbookKey, string folderPath)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _pendingPostCloseQueue.Enqueue(new PostCloseFollowUpRequest(workbookKey, folderPath, PostCloseRetryCount));
            if (_postClosePosted)
            {
                return;
            }

            _postClosePosted = true;
            EnsureDispatcher().BeginInvoke((MethodInvoker)ExecutePendingPostCloseQueue);
        }

        private void ExecutePendingPostCloseQueue()
        {
            _postClosePosted = false;

            while (_pendingPostCloseQueue.Count > 0)
            {
                PostCloseFollowUpRequest request = _pendingPostCloseQueue.Dequeue();
                if (request == null)
                {
                    continue;
                }

                try
                {
                    if (IsWorkbookStillOpen(request.WorkbookKey))
                    {
                        _logger.Info("Case workbook post-close follow-up skipped because workbook is still open. workbook=" + request.WorkbookKey);
                        continue;
                    }

                    QuitExcelIfNoVisibleWorkbook();
                }
                catch (COMException ex) when (ex.ErrorCode == ExcelBusyHResult && request.AttemptsRemaining > 0)
                {
                    _logger.Info(
                        "Case workbook post-close follow-up will retry because Excel is busy. workbook="
                        + request.WorkbookKey
                        + ", attemptsRemaining="
                        + request.AttemptsRemaining.ToString());
                    _pendingPostCloseQueue.Enqueue(request.NextAttempt());
                    SchedulePendingPostCloseRetry();
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error("Case workbook post-close follow-up failed.", ex);
                }
            }
        }

        private void SchedulePendingPostCloseRetry()
        {
            if (_postClosePosted)
            {
                return;
            }

            _postClosePosted = true;
            Timer retryTimer = new Timer();
            retryTimer.Interval = PostCloseRetryIntervalMs;
            retryTimer.Tick += (sender, args) =>
            {
                retryTimer.Stop();
                retryTimer.Dispose();
                ExecutePendingPostCloseQueue();
            };
            retryTimer.Start();
        }

        private void QuitExcelIfNoVisibleWorkbook()
        {
            bool hasVisibleWorkbook = false;
            foreach (Excel.Workbook openWorkbook in _application.Workbooks)
            {
                if (openWorkbook == null)
                {
                    continue;
                }

                try
                {
                    if (openWorkbook.Windows.Count > 0 && openWorkbook.Windows.Cast<Excel.Window>().Any(window => window.Visible))
                    {
                        hasVisibleWorkbook = true;
                        break;
                    }
                }
                catch
                {
                    // Closing workbook may already be tearing down. Ignore and keep scanning.
                }
            }

            _logger.Info("Case post-close visible workbook check. hasVisibleWorkbook=" + hasVisibleWorkbook.ToString());
            if (hasVisibleWorkbook)
            {
                return;
            }

            _logger.Info("Case post-close quitting Excel because no visible workbook remains.");
            bool previousDisplayAlerts = true;
            bool hasDisplayAlertsSnapshot = false;
            try
            {
                previousDisplayAlerts = _application.DisplayAlerts;
                hasDisplayAlertsSnapshot = true;
                _application.DisplayAlerts = false;
                _application.Quit();
            }
            catch
            {
                if (hasDisplayAlertsSnapshot)
                {
                    try
                    {
                        _application.DisplayAlerts = previousDisplayAlerts;
                    }
                    catch
                    {
                    }
                }

                throw;
            }
        }

        private Control EnsureDispatcher()
        {
            if (_dispatcher != null && !_dispatcher.IsDisposed)
            {
                return _dispatcher;
            }

            _dispatcher = new Control();
            IntPtr unusedHandle = _dispatcher.Handle;
            return _dispatcher;
        }

        private bool IsWorkbookStillOpen(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return false;
            }

            foreach (Excel.Workbook openWorkbook in _application.Workbooks)
            {
                if (openWorkbook == null)
                {
                    continue;
                }

                if (string.Equals(GetWorkbookKey(openWorkbook), workbookKey, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetWorkbookKey(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            string fullName = _excelInteropService.GetWorkbookFullName(workbook);
            return string.IsNullOrWhiteSpace(fullName)
                ? _excelInteropService.GetWorkbookName(workbook)
                : fullName;
        }

        private sealed class PostCloseFollowUpRequest
        {
            internal PostCloseFollowUpRequest(string workbookKey, string folderPath, int attemptsRemaining)
            {
                WorkbookKey = workbookKey ?? string.Empty;
                FolderPath = folderPath ?? string.Empty;
                AttemptsRemaining = attemptsRemaining;
            }

            internal string WorkbookKey { get; }

            internal string FolderPath { get; }

            internal int AttemptsRemaining { get; }

            internal PostCloseFollowUpRequest NextAttempt()
            {
                return new PostCloseFollowUpRequest(WorkbookKey, FolderPath, AttemptsRemaining - 1);
            }
        }
    }
}
