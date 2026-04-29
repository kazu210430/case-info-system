using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelHomeCoordinator
    {
        private readonly ThisAddIn _addin;
        private readonly KernelHomeCasePaneSuppressionCoordinator _suppressionCoordinator;

        internal KernelHomeCoordinator(ThisAddIn addin, KernelHomeCasePaneSuppressionCoordinator suppressionCoordinator)
        {
            _addin = addin;
            _suppressionCoordinator = suppressionCoordinator;
        }

        internal void SuppressUpcomingKernelHomeDisplay(string reason, bool suppressOnOpen, bool suppressOnActivate)
        {
            _suppressionCoordinator.SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen, suppressOnActivate);
        }

        internal bool ShouldSuppressKernelHomeDisplay(string eventName)
        {
            return IsKernelHomeSuppressionActive(eventName, consume: true);
        }

        internal void HandleKernelWorkbookBecameAvailable(string eventName, Excel.Workbook workbook)
        {
            _addin.HandleKernelWorkbookBecameAvailable(eventName, workbook);
        }

        internal bool ShouldAutoShowKernelHomeForEvent(string eventName, Excel.Workbook workbook)
        {
            bool shouldShow = _addin.ShouldShowKernelHomeOnStartup(workbook);
            _addin.Logger.Info(
                "Kernel HOME display context evaluated. eventName="
                + (eventName ?? string.Empty)
                + ", shouldShow="
                + shouldShow.ToString()
                + ", workbook="
                + _addin.GetWorkbookFullNameForLogging(workbook));
            return shouldShow;
        }

        internal bool ShouldReloadVisibleKernelHomeForEvent(string eventName, Excel.Workbook workbook)
        {
            return _addin.IsKernelWorkbook(workbook);
        }

        internal bool IsKernelHomeSuppressionActive(string eventName, bool consume)
        {
            return _suppressionCoordinator.IsKernelHomeSuppressionActive(eventName, consume);
        }
    }
}
