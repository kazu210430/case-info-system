using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookAvailabilityService
    {
        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly ExcelInteropService _excelInteropService;
        private readonly KernelHomeCoordinator _kernelHomeCoordinator;
        private readonly Action<string> _showKernelHomePlaceholderWithExternalWorkbookSuppression;
        private readonly Logger _logger;

        internal KernelWorkbookAvailabilityService(
            KernelWorkbookService kernelWorkbookService,
            ExcelInteropService excelInteropService,
            KernelHomeCoordinator kernelHomeCoordinator,
            Action<string> showKernelHomePlaceholderWithExternalWorkbookSuppression,
            Logger logger)
        {
            _kernelWorkbookService = kernelWorkbookService;
            _excelInteropService = excelInteropService;
            _kernelHomeCoordinator = kernelHomeCoordinator;
            _showKernelHomePlaceholderWithExternalWorkbookSuppression = showKernelHomePlaceholderWithExternalWorkbookSuppression;
            _logger = logger;
        }

        internal void Handle(string eventName, Excel.Workbook workbook, KernelHomeForm kernelHomeForm)
        {
            try
            {
                if (_kernelWorkbookService == null)
                {
                    return;
                }

                if (workbook != null && !_kernelWorkbookService.IsKernelWorkbook(workbook))
                {
                    _logger.Info("HandleKernelWorkbookBecameAvailable skipped for non-kernel workbook. eventName=" + (eventName ?? string.Empty) + ", workbook=" + _excelInteropService.GetWorkbookFullName(workbook));
                    return;
                }

                KernelHomeDisplayAvailabilityState state = BuildDisplayAvailabilityState(eventName, workbook, kernelHomeForm);
                KernelHomeDisplayAction action = KernelHomeDisplayAvailabilityPolicy.Decide(
                    state.HasKernelWorkbookReached,
                    state.IsDisplayReady,
                    state.HasVisibleKernelHome,
                    state.IsSuppressed,
                    state.IsDisplayContextAllowed,
                    state.ShouldReloadVisibleKernelHome);
                _logger?.Info(
                    "Kernel HOME availability state. eventName="
                    + state.EventName
                    + ", hasKernelWorkbookReached="
                    + state.HasKernelWorkbookReached.ToString()
                    + ", isDisplayReady="
                    + state.IsDisplayReady.ToString()
                    + ", hasVisibleKernelHome="
                    + state.HasVisibleKernelHome.ToString()
                    + ", isSuppressed="
                    + state.IsSuppressed.ToString()
                    + ", isDisplayContextAllowed="
                    + state.IsDisplayContextAllowed.ToString()
                    + ", shouldReloadVisibleKernelHome="
                    + state.ShouldReloadVisibleKernelHome.ToString()
                    + ", action="
                    + action.ToString()
                    + ", workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));

                if (action == KernelHomeDisplayAction.Show)
                {
                    _logger.Info("Kernel HOME requested from " + state.EventName);
                    _kernelWorkbookService.ClearHomeWorkbookBinding("KernelWorkbookAvailabilityService.Show." + state.EventName);
                    _showKernelHomePlaceholderWithExternalWorkbookSuppression("HandleKernelWorkbookBecameAvailable." + state.EventName);
                    return;
                }

                if (action == KernelHomeDisplayAction.ReloadVisible)
                {
                    kernelHomeForm.ReloadSettings();
                    _logger.Info("KernelHomeForm reloaded after " + state.EventName);
                }
            }
            catch (Exception ex)
            {
                _logger.Error("HandleKernelWorkbookBecameAvailable failed. eventName=" + (eventName ?? string.Empty), ex);
            }
        }

        private KernelHomeDisplayAvailabilityState BuildDisplayAvailabilityState(string eventName, Excel.Workbook workbook, KernelHomeForm kernelHomeForm)
        {
            string normalizedEventName = eventName ?? string.Empty;
            bool hasVisibleKernelHome = kernelHomeForm != null && !kernelHomeForm.IsDisposed && kernelHomeForm.Visible;
            return new KernelHomeDisplayAvailabilityState
            {
                EventName = normalizedEventName,
                HasKernelWorkbookReached = workbook != null && _kernelWorkbookService.IsKernelWorkbook(workbook),
                IsDisplayReady = kernelHomeForm == null || kernelHomeForm.IsDisposed || !kernelHomeForm.Visible,
                HasVisibleKernelHome = hasVisibleKernelHome,
                IsSuppressed = _kernelHomeCoordinator.ShouldSuppressKernelHomeDisplay(normalizedEventName),
                IsDisplayContextAllowed = _kernelHomeCoordinator.ShouldAutoShowKernelHomeForEvent(normalizedEventName, workbook),
                ShouldReloadVisibleKernelHome = hasVisibleKernelHome && _kernelHomeCoordinator.ShouldReloadVisibleKernelHomeForEvent(normalizedEventName, workbook)
            };
        }

        private sealed class KernelHomeDisplayAvailabilityState
        {
            internal string EventName { get; set; } = string.Empty;

            internal bool HasKernelWorkbookReached { get; set; }

            internal bool IsDisplayReady { get; set; }

            internal bool HasVisibleKernelHome { get; set; }

            internal bool IsSuppressed { get; set; }

            internal bool IsDisplayContextAllowed { get; set; }

            internal bool ShouldReloadVisibleKernelHome { get; set; }
        }
    }

    internal enum KernelHomeDisplayAction
    {
        None = 0,
        Show = 1,
        ReloadVisible = 2
    }

    internal static class KernelHomeDisplayAvailabilityPolicy
    {
        internal static KernelHomeDisplayAction Decide(
            bool hasKernelWorkbookReached,
            bool isDisplayReady,
            bool hasVisibleKernelHome,
            bool isSuppressed,
            bool isDisplayContextAllowed,
            bool shouldReloadVisibleKernelHome)
        {
            if (!hasKernelWorkbookReached || isSuppressed)
            {
                return KernelHomeDisplayAction.None;
            }

            if (hasVisibleKernelHome && shouldReloadVisibleKernelHome)
            {
                return KernelHomeDisplayAction.ReloadVisible;
            }

            if (!isDisplayReady || !isDisplayContextAllowed)
            {
                return KernelHomeDisplayAction.None;
            }

            return KernelHomeDisplayAction.Show;
        }
    }
}
