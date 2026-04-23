using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class ExcelWindowRecoveryService
    {
        private static readonly IntPtr HwndTopMost = new IntPtr(-1);
        private static readonly IntPtr HwndNoTopMost = new IntPtr(-2);
        private const int SwShow = 5;
        private const int SwRestore = 9;
        private const uint SwpNoMove = 0x0002;
        private const uint SwpNoSize = 0x0001;
        private const uint SwpShowWindow = 0x0040;

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        /// <summary>
        internal ExcelWindowRecoveryService(Excel.Application application, ExcelInteropService excelInteropService, Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal bool TryRecoverWorkbookWindow(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowInternal(
                workbook,
                reason,
                bringToFront,
                ensureWindowVisible: true,
                activateWindow: true);
        }

        internal bool TryRecoverWorkbookWindowWithoutShowing(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowInternal(
                workbook,
                reason,
                bringToFront,
                ensureWindowVisible: false,
                activateWindow: false);
        }

        private bool TryRecoverWorkbookWindowInternal(Excel.Workbook workbook, string reason, bool bringToFront, bool ensureWindowVisible, bool activateWindow)
        {
            if (workbook == null)
            {
                return false;
            }

            string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
            bool recoveredScreenUpdating = EnsureScreenUpdatingEnabled(reason, workbookFullName);
            Excel.Window window = ResolveWindow(workbook);
            if (window == null)
            {
                _logger.Warn(
                    "Excel window recovery skipped because workbook window could not be resolved. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + workbookFullName);
                return false;
            }

            bool recoveredWindowVisibility = ensureWindowVisible
                && EnsureWindowVisible(window, reason, workbookFullName);
            bool recoveredWindowState = EnsureWindowRestored(window, reason, workbookFullName);
            bool recoveredApplicationVisibility = EnsureApplicationVisible(reason, workbookFullName);

            if (activateWindow)
            {
                try
                {
                    window.Activate();
                }
                catch (Exception ex)
                {
                    _logger.Debug(nameof(ExcelWindowRecoveryService), "window.Activate failed but recovery continues. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", message=" + ex.Message);
                }
            }

            if (bringToFront
                && (ensureWindowVisible || window.Visible)
                && (recoveredApplicationVisibility || recoveredScreenUpdating || recoveredWindowVisibility || recoveredWindowState))
            {
                PromoteExcelWindow(window, reason, workbookFullName);
            }

            _logger.Info(
                "Excel window recovery evaluated. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + workbookFullName
                + ", appRecovered="
                + recoveredApplicationVisibility.ToString()
                + ", screenUpdatingRecovered="
                + recoveredScreenUpdating.ToString()
                + ", windowVisibleRecovered="
                + recoveredWindowVisibility.ToString()
                + ", windowStateRecovered="
                + recoveredWindowState.ToString()
                + ", ensureWindowVisible="
                + ensureWindowVisible.ToString()
                + ", activateWindow="
                + activateWindow.ToString());
            return true;
        }

        /// <summary>
        internal bool TryRecoverActiveWorkbookWindow(string reason, bool bringToFront)
        {
            Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();
            return activeWorkbook != null && TryRecoverWorkbookWindow(activeWorkbook, reason, bringToFront);
        }

        internal bool TryRecoverActiveWorkbookWindowWithoutShowing(string reason, bool bringToFront)
        {
            Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();
            return activeWorkbook != null && TryRecoverWorkbookWindowWithoutShowing(activeWorkbook, reason, bringToFront);
        }

        /// <summary>
        internal bool EnsureApplicationVisible(string reason, string workbookFullName)
        {
            try
            {
                if (_application.Visible)
                {
                    return false;
                }

                _application.Visible = true;
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                ShowWindow(hwnd, SwRestore);
                ShowWindow(hwnd, SwShow);
                _logger.Info(
                    "Excel application visibility recovered. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("EnsureApplicationVisible failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        /// <summary>
        internal bool EnsureScreenUpdatingEnabled(string reason, string workbookFullName)
        {
            try
            {
                if (_application.ScreenUpdating)
                {
                    return false;
                }

                _application.ScreenUpdating = true;
                _logger.Info(
                    "Excel screen updating recovered. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("EnsureScreenUpdatingEnabled failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        /// <summary>
        private Excel.Window ResolveWindow(Excel.Workbook workbook)
        {
            Excel.Window visibleWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
            if (visibleWindow != null)
            {
                return visibleWindow;
            }

            try
            {
                return workbook.Windows.Count > 0 ? workbook.Windows[1] : null;
            }
            catch (Exception ex)
            {
                _logger.Error("ResolveWindow failed.", ex);
                return null;
            }
        }

        /// <summary>
        private bool EnsureWindowVisible(Excel.Window window, string reason, string workbookFullName)
        {
            try
            {
                if (window.Visible)
                {
                    return false;
                }

                window.Visible = true;
                _logger.Info(
                    "Workbook window visibility recovered. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("EnsureWindowVisible failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        private bool EnsureWindowRestored(Excel.Window window, string reason, string workbookFullName)
        {
            try
            {
                if (window.WindowState != Excel.XlWindowState.xlMinimized)
                {
                    return false;
                }

                window.WindowState = Excel.XlWindowState.xlNormal;
                _logger.Info(
                    "Workbook window state recovered. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("EnsureWindowRestored failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        /// <summary>
        private void PromoteExcelWindow(Excel.Window window, string reason, string workbookFullName)
        {
            try
            {
                PromoteWindow(new IntPtr(window.Hwnd));
                _logger.Info(
                    "Excel window promoted after recovery. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
            }
            catch (Exception ex)
            {
                _logger.Debug(nameof(ExcelWindowRecoveryService), "PromoteWindow failed but recovery continues. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty) + ", message=" + ex.Message);
            }
        }

        /// <summary>
        private static void PromoteWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return;
            }

            ShowWindow(hwnd, SwRestore);
            SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpShowWindow);
            SetWindowPos(hwnd, HwndNoTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpShowWindow);
            SetForegroundWindow(hwnd);
        }
    }
}
