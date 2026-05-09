using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneReadyShowRetryScheduler
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Logger _logger;
        private readonly TaskPaneRetryTimerLifecycle _retryTimerLifecycle;
        private readonly Func<Excel.Workbook, string> _formatWorkbookDescriptor;
        private readonly Func<Excel.Workbook, string> _safeWorkbookFullName;
        private readonly int _retryDelayMs;
        private readonly int _maxAttempts;

        internal TaskPaneReadyShowRetryScheduler(
            Logger logger,
            TaskPaneRetryTimerLifecycle retryTimerLifecycle,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<Excel.Workbook, string> safeWorkbookFullName)
            : this(
                logger,
                retryTimerLifecycle,
                formatWorkbookDescriptor,
                safeWorkbookFullName,
                WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowRetryDelayMs,
                WorkbookTaskPaneReadyShowAttemptWorker.ReadyShowMaxAttempts)
        {
        }

        internal TaskPaneReadyShowRetryScheduler(
            Logger logger,
            TaskPaneRetryTimerLifecycle retryTimerLifecycle,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<Excel.Workbook, string> safeWorkbookFullName,
            int retryDelayMs,
            int maxAttempts)
        {
            _logger = logger;
            _retryTimerLifecycle = retryTimerLifecycle ?? throw new ArgumentNullException(nameof(retryTimerLifecycle));
            _formatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
            _safeWorkbookFullName = safeWorkbookFullName ?? throw new ArgumentNullException(nameof(safeWorkbookFullName));
            _retryDelayMs = retryDelayMs;
            _maxAttempts = maxAttempts;
        }

        internal void Schedule(Excel.Workbook workbook, string reason, int attemptNumber, Action retryAction)
        {
            string safeReason = reason ?? string.Empty;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneReadyShowRetryScheduler action=wait-ready-retry-scheduled reason="
                + safeReason
                + ", readyShowReason="
                + safeReason
                + ", workbook="
                + _formatWorkbookDescriptor(workbook)
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + _maxAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryScheduled=true"
                + ", retryDelayMs="
                + _retryDelayMs.ToString(CultureInfo.InvariantCulture)
                + ", delayMs="
                + _retryDelayMs.ToString(CultureInfo.InvariantCulture));
            _logger?.Info(
                "TaskPane wait-ready retry scheduled. reason="
                + safeReason
                + ", workbook="
                + _safeWorkbookFullName(workbook)
                + ", readyShowReason="
                + safeReason
                + ", attempt="
                + attemptNumber.ToString(CultureInfo.InvariantCulture)
                + ", maxAttempts="
                + _maxAttempts.ToString(CultureInfo.InvariantCulture)
                + ", retryScheduled=true"
                + ", retryDelayMs="
                + _retryDelayMs.ToString(CultureInfo.InvariantCulture));

            if (retryAction == null)
            {
                return;
            }

            _retryTimerLifecycle.ScheduleWaitReadyRetryTimer(
                _retryDelayMs,
                () =>
                {
                    _logger?.Info(
                        KernelFlickerTracePrefix
                        + " source=TaskPaneReadyShowRetryScheduler action=wait-ready-retry-firing reason="
                        + safeReason
                        + ", readyShowReason="
                        + safeReason
                        + ", workbook="
                        + _formatWorkbookDescriptor(workbook)
                        + ", attempt="
                        + attemptNumber.ToString(CultureInfo.InvariantCulture)
                        + ", maxAttempts="
                        + _maxAttempts.ToString(CultureInfo.InvariantCulture)
                        + ", retryDelayMs="
                        + _retryDelayMs.ToString(CultureInfo.InvariantCulture));
                    _logger?.Info(
                        "TaskPane wait-ready retry firing. reason="
                        + safeReason
                        + ", workbook="
                        + _safeWorkbookFullName(workbook)
                        + ", readyShowReason="
                        + safeReason
                        + ", attempt="
                        + attemptNumber.ToString(CultureInfo.InvariantCulture)
                        + ", maxAttempts="
                        + _maxAttempts.ToString(CultureInfo.InvariantCulture)
                        + ", retryDelayMs="
                        + _retryDelayMs.ToString(CultureInfo.InvariantCulture));
                    retryAction();
                });
        }
    }
}
