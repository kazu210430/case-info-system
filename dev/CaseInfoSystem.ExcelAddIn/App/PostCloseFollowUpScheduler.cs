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
        private const int ExcelBusyRetryCount = 20;
        private const int ExcelBusyRetryIntervalMs = 500;
        private const int TargetStillOpenRetryCount = 5;
        private const int TargetStillOpenRetryIntervalMs = 250;
        private const string WhiteExcelPreventionQueued = "WhiteExcelPreventionQueued";
        private const string WhiteExcelPreventionNotRequired = "WhiteExcelPreventionNotRequired";
        private const string WhiteExcelPreventionCompleted = "WhiteExcelPreventionCompleted";
        private const string WhiteExcelPreventionFailed = "WhiteExcelPreventionFailed";

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;
        private readonly ManagedWorkbookCloseMarkerStore _managedCloseMarkerStore;
        private readonly Queue<PostCloseFollowUpRequest> _pendingPostCloseQueue = new Queue<PostCloseFollowUpRequest>();
        private Control _dispatcher;
        private bool _postClosePosted;

        internal PostCloseFollowUpScheduler(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            Logger logger)
            : this(application, excelInteropService, logger, null)
        {
        }

        internal PostCloseFollowUpScheduler(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            Logger logger,
            ManagedWorkbookCloseMarkerStore managedCloseMarkerStore)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _managedCloseMarkerStore = managedCloseMarkerStore;
        }

        internal void Schedule(string workbookKey, string folderPath)
        {
            ScheduleManagedWorkbookClose(workbookKey, folderPath, ManagedWorkbookCloseMarkerKind.CaseClose);
        }

        internal void ScheduleManagedWorkbookClose(string workbookKey, string folderPath, ManagedWorkbookCloseMarkerKind closeKind)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            TryWriteManagedCloseMarker(closeKind, workbookKey);
            PostCloseFollowUpRequest queuedRequest = new PostCloseFollowUpRequest(
                workbookKey,
                folderPath,
                attemptNumber: 1,
                targetStillOpenRetriesRemaining: TargetStillOpenRetryCount,
                excelBusyRetriesRemaining: ExcelBusyRetryCount,
                closeKind: closeKind);
            _pendingPostCloseQueue.Enqueue(queuedRequest);
            LogWhiteExcelPreventionOutcome(
                WhiteExcelPreventionQueued,
                workbookKey,
                hasVisibleWorkbook: null,
                quitAttempted: false,
                quitCompleted: false,
                reason: "postCloseFollowUpQueued",
                pendingQueueCount: _pendingPostCloseQueue.Count,
                attemptsRemaining: queuedRequest.ExcelBusyRetriesRemaining,
                folderPathPresent: !string.IsNullOrWhiteSpace(queuedRequest.FolderPath),
                targetWorkbookStillOpen: null,
                closeKind: closeKind,
                attemptNumber: queuedRequest.AttemptNumber,
                targetStillOpenRetriesRemaining: queuedRequest.TargetStillOpenRetriesRemaining,
                excelBusyRetriesRemaining: queuedRequest.ExcelBusyRetriesRemaining);
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

            bool retryRequested = false;
            int retryIntervalMs = TargetStillOpenRetryIntervalMs;
            string retryReason = string.Empty;
            int requestsToProcess = _pendingPostCloseQueue.Count;

            while (requestsToProcess > 0 && _pendingPostCloseQueue.Count > 0)
            {
                requestsToProcess--;
                PostCloseFollowUpRequest request = _pendingPostCloseQueue.Dequeue();
                if (request == null)
                {
                    continue;
                }

                try
                {
                    PostCloseApplicationFacts applicationFacts = CapturePostCloseApplicationFacts();
                    _logger.Info(
                        "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                        + " action=post-close-follow-up-request-dequeued"
                        + " workbook=" + request.WorkbookKey
                        + ", managedCloseKind=" + request.CloseKind.ToString()
                        + ", pendingQueueCount=" + _pendingPostCloseQueue.Count.ToString()
                        + ", attemptsRemaining=" + request.ExcelBusyRetriesRemaining.ToString()
                        + ", attemptNumber=" + request.AttemptNumber.ToString()
                        + ", targetStillOpenRetriesRemaining=" + request.TargetStillOpenRetriesRemaining.ToString()
                        + ", excelBusyRetriesRemaining=" + request.ExcelBusyRetriesRemaining.ToString()
                        + ", folderPathPresent=" + (!string.IsNullOrWhiteSpace(request.FolderPath)).ToString()
                        + applicationFacts.ToTraceFields());
                    bool targetWorkbookStillOpen = IsWorkbookStillOpen(request.WorkbookKey);
                    string decision = GetPostCloseDecision(targetWorkbookStillOpen, request);
                    _logger.Info(
                        "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                        + " action=post-close-follow-up-decision"
                        + " workbook=" + request.WorkbookKey
                        + ", managedCloseKind=" + request.CloseKind.ToString()
                        + ", targetWorkbookStillOpen=" + targetWorkbookStillOpen.ToString()
                        + ", pendingQueueCount=" + _pendingPostCloseQueue.Count.ToString()
                        + ", attemptsRemaining=" + request.ExcelBusyRetriesRemaining.ToString()
                        + ", attemptNumber=" + request.AttemptNumber.ToString()
                        + ", targetStillOpenRetriesRemaining=" + request.TargetStillOpenRetriesRemaining.ToString()
                        + ", excelBusyRetriesRemaining=" + request.ExcelBusyRetriesRemaining.ToString()
                        + ", decision=" + decision
                        + applicationFacts.ToTraceFields());
                    if (targetWorkbookStillOpen)
                    {
                        if (request.TargetStillOpenRetriesRemaining > 0)
                        {
                            PostCloseFollowUpRequest retryRequest = request.NextTargetStillOpenAttempt();
                            _pendingPostCloseQueue.Enqueue(retryRequest);
                            retryRequested = true;
                            retryIntervalMs = TargetStillOpenRetryIntervalMs;
                            retryReason = "targetWorkbookStillOpen";
                            LogPostCloseRetryScheduled(
                                request,
                                retryRequest,
                                retryReason,
                                TargetStillOpenRetryIntervalMs,
                                applicationFacts,
                                targetWorkbookStillOpen);
                            continue;
                        }

                        LogWhiteExcelPreventionOutcome(
                            WhiteExcelPreventionNotRequired,
                            request.WorkbookKey,
                            hasVisibleWorkbook: null,
                            quitAttempted: false,
                            quitCompleted: false,
                            reason: "targetWorkbookStillOpenRetryExhausted",
                            pendingQueueCount: _pendingPostCloseQueue.Count,
                            attemptsRemaining: request.ExcelBusyRetriesRemaining,
                            folderPathPresent: !string.IsNullOrWhiteSpace(request.FolderPath),
                            targetWorkbookStillOpen: true,
                            closeKind: request.CloseKind,
                            attemptNumber: request.AttemptNumber,
                            targetStillOpenRetriesRemaining: request.TargetStillOpenRetriesRemaining,
                            excelBusyRetriesRemaining: request.ExcelBusyRetriesRemaining,
                            applicationFacts: applicationFacts);
                        _logger.Info(
                            "Case workbook post-close follow-up skipped because workbook is still open after retry exhaustion. workbook="
                            + request.WorkbookKey);
                        continue;
                    }

                    QuitExcelIfNoVisibleWorkbook(request.CloseKind, request, applicationFacts);
                }
                catch (COMException ex) when (ex.ErrorCode == ExcelBusyHResult && request.ExcelBusyRetriesRemaining > 0)
                {
                    PostCloseFollowUpRequest retryRequest = request.NextExcelBusyAttempt();
                    _logger.Info(
                        "Case workbook post-close follow-up will retry because Excel is busy. workbook="
                        + request.WorkbookKey
                        + ", attemptsRemaining="
                        + request.ExcelBusyRetriesRemaining.ToString()
                        + ", attemptNumber="
                        + request.AttemptNumber.ToString()
                        + ", nextAttemptNumber="
                        + retryRequest.AttemptNumber.ToString());
                    _pendingPostCloseQueue.Enqueue(retryRequest);
                    SchedulePendingPostCloseRetry(ExcelBusyRetryIntervalMs, "excelBusy");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error("Case workbook post-close follow-up failed.", ex);
                }
            }

            if (retryRequested)
            {
                SchedulePendingPostCloseRetry(retryIntervalMs, retryReason);
            }
        }

        private void SchedulePendingPostCloseRetry(int retryIntervalMs, string reason)
        {
            if (_postClosePosted)
            {
                return;
            }

            _postClosePosted = true;
            Timer retryTimer = new Timer();
            retryTimer.Interval = retryIntervalMs <= 0 ? TargetStillOpenRetryIntervalMs : retryIntervalMs;
            retryTimer.Tick += (sender, args) =>
            {
                retryTimer.Stop();
                retryTimer.Dispose();
                ExecutePendingPostCloseQueue();
            };
            retryTimer.Start();
            _logger.Info(
                "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                + " action=post-close-follow-up-retry-timer-scheduled"
                + " retryReason=" + (reason ?? string.Empty)
                + ", retryDelayMs=" + retryTimer.Interval.ToString()
                + ", pendingQueueCount=" + _pendingPostCloseQueue.Count.ToString());
        }

        private void QuitExcelIfNoVisibleWorkbook()
        {
            QuitExcelIfNoVisibleWorkbook(ManagedWorkbookCloseMarkerKind.CaseClose);
        }

        private void QuitExcelIfNoVisibleWorkbook(ManagedWorkbookCloseMarkerKind closeKind)
        {
            QuitExcelIfNoVisibleWorkbook(closeKind, null, CapturePostCloseApplicationFacts());
        }

        private void QuitExcelIfNoVisibleWorkbook(
            ManagedWorkbookCloseMarkerKind closeKind,
            PostCloseFollowUpRequest request,
            PostCloseApplicationFacts applicationFacts)
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

            _logger.Info(
                "Case post-close visible workbook check. hasVisibleWorkbook="
                + hasVisibleWorkbook.ToString()
                + (applicationFacts == null ? string.Empty : applicationFacts.ToTraceFields()));
            if (hasVisibleWorkbook)
            {
                LogWhiteExcelPreventionOutcome(
                    WhiteExcelPreventionNotRequired,
                    workbookKey: request == null ? string.Empty : request.WorkbookKey,
                    hasVisibleWorkbook: true,
                    quitAttempted: false,
                    quitCompleted: false,
                    reason: "visibleWorkbookExists",
                    closeKind: closeKind,
                    attemptNumber: request == null ? (int?)null : request.AttemptNumber,
                    targetStillOpenRetriesRemaining: request == null ? (int?)null : request.TargetStillOpenRetriesRemaining,
                    excelBusyRetriesRemaining: request == null ? (int?)null : request.ExcelBusyRetriesRemaining,
                    applicationFacts: applicationFacts);
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
                    workbookKey: request == null ? string.Empty : request.WorkbookKey,
                    hasVisibleWorkbook: false,
                    quitAttempted: true,
                    quitCompleted: true,
                    reason: "noVisibleWorkbookQuitCompleted",
                    closeKind: closeKind,
                    attemptNumber: request == null ? (int?)null : request.AttemptNumber,
                    targetStillOpenRetriesRemaining: request == null ? (int?)null : request.TargetStillOpenRetriesRemaining,
                    excelBusyRetriesRemaining: request == null ? (int?)null : request.ExcelBusyRetriesRemaining,
                    applicationFacts: applicationFacts);
            }
            catch
            {
                LogWhiteExcelPreventionOutcome(
                    WhiteExcelPreventionFailed,
                    workbookKey: request == null ? string.Empty : request.WorkbookKey,
                    hasVisibleWorkbook: false,
                    quitAttempted: true,
                    quitCompleted: false,
                    reason: "quitFailed",
                    closeKind: closeKind,
                    attemptNumber: request == null ? (int?)null : request.AttemptNumber,
                    targetStillOpenRetriesRemaining: request == null ? (int?)null : request.TargetStillOpenRetriesRemaining,
                    excelBusyRetriesRemaining: request == null ? (int?)null : request.ExcelBusyRetriesRemaining,
                    applicationFacts: applicationFacts);
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
            bool? targetWorkbookStillOpen = null,
            ManagedWorkbookCloseMarkerKind? closeKind = null,
            int? attemptNumber = null,
            int? targetStillOpenRetriesRemaining = null,
            int? excelBusyRetriesRemaining = null,
            PostCloseApplicationFacts applicationFacts = null)
        {
            _logger.Info(
                "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                + " action=white-excel-prevention-outcome"
                + " whiteExcelPreventionOutcome=" + (outcome ?? string.Empty)
                + ", managedCloseKind=" + (closeKind.HasValue ? closeKind.Value.ToString() : "unknown")
                + ", workbook=" + (workbookKey ?? string.Empty)
                + ", hasVisibleWorkbook=" + (hasVisibleWorkbook.HasValue ? hasVisibleWorkbook.Value.ToString() : "unknown")
                + ", quitAttempted=" + quitAttempted.ToString()
                + ", quitCompleted=" + quitCompleted.ToString()
                + ", outcomeReason=" + (reason ?? string.Empty)
                + ", pendingQueueCount=" + (pendingQueueCount.HasValue ? pendingQueueCount.Value.ToString() : "unknown")
                + ", attemptsRemaining=" + (attemptsRemaining.HasValue ? attemptsRemaining.Value.ToString() : "unknown")
                + ", folderPathPresent=" + (folderPathPresent.HasValue ? folderPathPresent.Value.ToString() : "unknown")
                + ", targetWorkbookStillOpen=" + (targetWorkbookStillOpen.HasValue ? targetWorkbookStillOpen.Value.ToString() : "unknown")
                + ", attemptNumber=" + (attemptNumber.HasValue ? attemptNumber.Value.ToString() : "unknown")
                + ", targetStillOpenRetriesRemaining=" + (targetStillOpenRetriesRemaining.HasValue ? targetStillOpenRetriesRemaining.Value.ToString() : "unknown")
                + ", excelBusyRetriesRemaining=" + (excelBusyRetriesRemaining.HasValue ? excelBusyRetriesRemaining.Value.ToString() : "unknown")
                + (applicationFacts == null ? string.Empty : applicationFacts.ToTraceFields()));
        }

        private static string GetPostCloseDecision(bool targetWorkbookStillOpen, PostCloseFollowUpRequest request)
        {
            if (!targetWorkbookStillOpen)
            {
                return "scan-visible-workbooks";
            }

            return request != null && request.TargetStillOpenRetriesRemaining > 0
                ? "retry-target-still-open"
                : "skip-still-open-retry-exhausted";
        }

        private void LogPostCloseRetryScheduled(
            PostCloseFollowUpRequest request,
            PostCloseFollowUpRequest retryRequest,
            string reason,
            int retryDelayMs,
            PostCloseApplicationFacts applicationFacts,
            bool targetWorkbookStillOpen)
        {
            _logger.Info(
                "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                + " action=post-close-follow-up-retry-scheduled"
                + " workbook=" + (request == null ? string.Empty : request.WorkbookKey)
                + ", managedCloseKind=" + (request == null ? "unknown" : request.CloseKind.ToString())
                + ", retryReason=" + (reason ?? string.Empty)
                + ", retryDelayMs=" + retryDelayMs.ToString()
                + ", targetWorkbookStillOpen=" + targetWorkbookStillOpen.ToString()
                + ", attemptNumber=" + (request == null ? "unknown" : request.AttemptNumber.ToString())
                + ", nextAttemptNumber=" + (retryRequest == null ? "unknown" : retryRequest.AttemptNumber.ToString())
                + ", targetStillOpenRetriesRemaining=" + (retryRequest == null ? "unknown" : retryRequest.TargetStillOpenRetriesRemaining.ToString())
                + ", excelBusyRetriesRemaining=" + (retryRequest == null ? "unknown" : retryRequest.ExcelBusyRetriesRemaining.ToString())
                + ", pendingQueueCount=" + _pendingPostCloseQueue.Count.ToString()
                + (applicationFacts == null ? string.Empty : applicationFacts.ToTraceFields()));
        }

        private PostCloseApplicationFacts CapturePostCloseApplicationFacts()
        {
            var facts = new PostCloseApplicationFacts();

            try
            {
                facts.ApplicationVisible = _application != null && _application.Visible;
            }
            catch
            {
                facts.ReadFailed = true;
                facts.ApplicationVisibleReadFailed = true;
            }

            try
            {
                Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
                facts.ActiveWorkbookPresent = activeWorkbook != null;
            }
            catch
            {
                facts.ReadFailed = true;
                facts.ActiveWorkbookReadFailed = true;
            }

            try
            {
                facts.WorkbooksCount = _application == null || _application.Workbooks == null ? -1 : _application.Workbooks.Count;
            }
            catch
            {
                facts.ReadFailed = true;
                facts.WorkbooksCount = -1;
                facts.WorkbooksCountReadFailed = true;
            }

            return facts;
        }

        private void TryWriteManagedCloseMarker(ManagedWorkbookCloseMarkerKind closeKind, string workbookKey)
        {
            if (_managedCloseMarkerStore == null)
            {
                return;
            }

            try
            {
                _managedCloseMarkerStore.Write(closeKind, workbookKey);
                _logger.Info(
                    "[KernelFlickerTrace] source=PostCloseFollowUpScheduler"
                    + " action=managed-close-marker-written"
                    + " managedCloseKind=" + closeKind.ToString()
                    + ", markerPath=" + _managedCloseMarkerStore.MarkerPath
                    + ", ttlSeconds=" + ManagedWorkbookCloseMarkerStore.DefaultTimeToLiveSeconds.ToString()
                    + ", workbook=" + (workbookKey ?? string.Empty));
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Managed close marker write failed. managedCloseKind="
                    + closeKind.ToString()
                    + ", workbook="
                    + (workbookKey ?? string.Empty),
                    ex);
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
            internal PostCloseFollowUpRequest(
                string workbookKey,
                string folderPath,
                int attemptNumber,
                int targetStillOpenRetriesRemaining,
                int excelBusyRetriesRemaining,
                ManagedWorkbookCloseMarkerKind closeKind)
            {
                WorkbookKey = workbookKey ?? string.Empty;
                FolderPath = folderPath ?? string.Empty;
                AttemptNumber = attemptNumber;
                TargetStillOpenRetriesRemaining = targetStillOpenRetriesRemaining;
                ExcelBusyRetriesRemaining = excelBusyRetriesRemaining;
                CloseKind = closeKind;
            }

            internal string WorkbookKey { get; }

            internal string FolderPath { get; }

            internal int AttemptNumber { get; }

            internal int TargetStillOpenRetriesRemaining { get; }

            internal int ExcelBusyRetriesRemaining { get; }

            internal ManagedWorkbookCloseMarkerKind CloseKind { get; }

            internal PostCloseFollowUpRequest NextTargetStillOpenAttempt()
            {
                return new PostCloseFollowUpRequest(
                    WorkbookKey,
                    FolderPath,
                    AttemptNumber + 1,
                    TargetStillOpenRetriesRemaining - 1,
                    ExcelBusyRetriesRemaining,
                    CloseKind);
            }

            internal PostCloseFollowUpRequest NextExcelBusyAttempt()
            {
                return new PostCloseFollowUpRequest(
                    WorkbookKey,
                    FolderPath,
                    AttemptNumber + 1,
                    TargetStillOpenRetriesRemaining,
                    ExcelBusyRetriesRemaining - 1,
                    CloseKind);
            }
        }

        private sealed class PostCloseApplicationFacts
        {
            internal int WorkbooksCount { get; set; } = -1;

            internal bool ActiveWorkbookPresent { get; set; }

            internal bool ApplicationVisible { get; set; }

            internal bool ReadFailed { get; set; }

            internal bool ApplicationVisibleReadFailed { get; set; }

            internal bool ActiveWorkbookReadFailed { get; set; }

            internal bool WorkbooksCountReadFailed { get; set; }

            internal string ToTraceFields()
            {
                return ", workbooksCount=" + WorkbooksCount.ToString()
                    + ", activeWorkbookPresent=" + ActiveWorkbookPresent.ToString()
                    + ", applicationVisible=" + ApplicationVisible.ToString()
                    + ", postCloseFactsReadFailed=" + ReadFailed.ToString()
                    + ", applicationVisibleReadFailed=" + ApplicationVisibleReadFailed.ToString()
                    + ", activeWorkbookReadFailed=" + ActiveWorkbookReadFailed.ToString()
                    + ", workbooksCountReadFailed=" + WorkbooksCountReadFailed.ToString();
            }
        }
    }
}
