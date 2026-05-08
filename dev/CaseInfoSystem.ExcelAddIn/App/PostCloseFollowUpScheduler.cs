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
        private const string WhiteExcelPreventionQueued = "WhiteExcelPreventionQueued";
        private const string WhiteExcelPreventionNotRequired = "WhiteExcelPreventionNotRequired";
        private const string WhiteExcelPreventionCompleted = "WhiteExcelPreventionCompleted";
        private const string WhiteExcelPreventionFailed = "WhiteExcelPreventionFailed";

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

            PostCloseFollowUpRequest queuedRequest = new PostCloseFollowUpRequest(workbookKey, folderPath, PostCloseRetryCount);
            _pendingPostCloseQueue.Enqueue(queuedRequest);
            LogWhiteExcelPreventionOutcome(
                WhiteExcelPreventionQueued,
                workbookKey,
                hasVisibleWorkbook: null,
                quitAttempted: false,
                quitCompleted: false,
                reason: "postCloseFollowUpQueued",
                pendingQueueCount: _pendingPostCloseQueue.Count,
                attemptsRemaining: queuedRequest.AttemptsRemaining,
                folderPathPresent: !string.IsNullOrWhiteSpace(queuedRequest.FolderPath),
                targetWorkbookStillOpen: null);
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
                    _logger.Info(
                        "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                        + " action=post-close-follow-up-request-dequeued"
                        + " workbook=" + request.WorkbookKey
                        + ", pendingQueueCount=" + _pendingPostCloseQueue.Count.ToString()
                        + ", attemptsRemaining=" + request.AttemptsRemaining.ToString()
                        + ", folderPathPresent=" + (!string.IsNullOrWhiteSpace(request.FolderPath)).ToString());
                    bool targetWorkbookStillOpen = IsWorkbookStillOpen(request.WorkbookKey);
                    _logger.Info(
                        "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                        + " action=post-close-follow-up-decision"
                        + " workbook=" + request.WorkbookKey
                        + ", targetWorkbookStillOpen=" + targetWorkbookStillOpen.ToString()
                        + ", pendingQueueCount=" + _pendingPostCloseQueue.Count.ToString()
                        + ", attemptsRemaining=" + request.AttemptsRemaining.ToString()
                        + ", decision=" + (targetWorkbookStillOpen ? "skip-still-open" : "scan-visible-workbooks"));
                    if (targetWorkbookStillOpen)
                    {
                        LogWhiteExcelPreventionOutcome(
                            WhiteExcelPreventionNotRequired,
                            request.WorkbookKey,
                            hasVisibleWorkbook: null,
                            quitAttempted: false,
                            quitCompleted: false,
                            reason: "targetWorkbookStillOpen",
                            pendingQueueCount: _pendingPostCloseQueue.Count,
                            attemptsRemaining: request.AttemptsRemaining,
                            folderPathPresent: !string.IsNullOrWhiteSpace(request.FolderPath),
                            targetWorkbookStillOpen: true);
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
                LogWhiteExcelPreventionOutcome(
                    WhiteExcelPreventionNotRequired,
                    workbookKey: string.Empty,
                    hasVisibleWorkbook: true,
                    quitAttempted: false,
                    quitCompleted: false,
                    reason: "visibleWorkbookExists");
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
                LogWhiteExcelPreventionOutcome(
                    WhiteExcelPreventionCompleted,
                    workbookKey: string.Empty,
                    hasVisibleWorkbook: false,
                    quitAttempted: true,
                    quitCompleted: true,
                    reason: "noVisibleWorkbookQuitCompleted");
            }
            catch
            {
                LogWhiteExcelPreventionOutcome(
                    WhiteExcelPreventionFailed,
                    workbookKey: string.Empty,
                    hasVisibleWorkbook: false,
                    quitAttempted: true,
                    quitCompleted: false,
                    reason: "quitFailed");
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

        private void LogWhiteExcelPreventionOutcome(
            string outcome,
            string workbookKey,
            bool? hasVisibleWorkbook,
            bool quitAttempted,
            bool quitCompleted,
            string reason,
            int? pendingQueueCount = null,
            int? attemptsRemaining = null,
            bool? folderPathPresent = null,
            bool? targetWorkbookStillOpen = null)
        {
            _logger.Info(
                "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                + " action=white-excel-prevention-outcome"
                + " whiteExcelPreventionOutcome=" + (outcome ?? string.Empty)
                + ", workbook=" + (workbookKey ?? string.Empty)
                + ", hasVisibleWorkbook=" + (hasVisibleWorkbook.HasValue ? hasVisibleWorkbook.Value.ToString() : "unknown")
                + ", quitAttempted=" + quitAttempted.ToString()
                + ", quitCompleted=" + quitCompleted.ToString()
                + ", outcomeReason=" + (reason ?? string.Empty)
                + ", pendingQueueCount=" + (pendingQueueCount.HasValue ? pendingQueueCount.Value.ToString() : "unknown")
                + ", attemptsRemaining=" + (attemptsRemaining.HasValue ? attemptsRemaining.Value.ToString() : "unknown")
                + ", folderPathPresent=" + (folderPathPresent.HasValue ? folderPathPresent.Value.ToString() : "unknown")
                + ", targetWorkbookStillOpen=" + (targetWorkbookStillOpen.HasValue ? targetWorkbookStillOpen.Value.ToString() : "unknown"));
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
