using System;
using System.Globalization;
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
                activateWindow: true,
                allowWindowCreation: true);
        }

        internal bool TryRecoverWorkbookWindowUsingExistingWindows(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowInternal(
                workbook,
                reason,
                bringToFront,
                ensureWindowVisible: true,
                activateWindow: true,
                allowWindowCreation: false);
        }

        internal bool TryRecoverWorkbookWindowWithoutShowing(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowInternal(
                workbook,
                reason,
                bringToFront,
                ensureWindowVisible: false,
                activateWindow: false,
                allowWindowCreation: true);
        }

        internal bool TryRecoverWorkbookWindowWithoutShowingUsingExistingWindows(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowInternal(
                workbook,
                reason,
                bringToFront,
                ensureWindowVisible: false,
                activateWindow: false,
                allowWindowCreation: false);
        }

        private bool TryRecoverWorkbookWindowInternal(Excel.Workbook workbook, string reason, bool bringToFront, bool ensureWindowVisible, bool activateWindow, bool allowWindowCreation)
        {
            if (workbook == null)
            {
                return false;
            }

            string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
            bool recoveredScreenUpdating = EnsureScreenUpdatingEnabled(reason, workbookFullName);
            LogWindowResolutionSnapshot("before-resolve", workbook, reason, workbookFullName, allowWindowCreation);
            Excel.Window window = ResolveWindow(workbook, reason, workbookFullName, allowWindowCreation);
            if (window == null)
            {
                _logger.Warn(
                    "Excel window recovery skipped because workbook window could not be resolved. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + workbookFullName
                    + ", allowWindowCreation="
                    + allowWindowCreation.ToString());
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
                + activateWindow.ToString()
                + ", allowWindowCreation="
                + allowWindowCreation.ToString()
                + ", resolvedWindow="
                + DescribeWindow(window));
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
        private Excel.Window ResolveWindow(Excel.Workbook workbook, string reason, string workbookFullName, bool allowWindowCreation)
        {
            Excel.Window visibleWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
            if (visibleWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook visible window. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(visibleWindow));
                return visibleWindow;
            }

            Excel.Window workbookWindow = TryGetFirstWorkbookWindow(workbook);
            if (workbookWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook window from Workbook.Windows. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(workbookWindow));
                return workbookWindow;
            }

            Excel.Window applicationWindow = TryFindApplicationWindowForWorkbook(workbook, workbookFullName, visibleOnly: true);
            if (applicationWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook window from Application.Windows. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(applicationWindow));
                return applicationWindow;
            }

            Excel.Window activeWindow = TryGetActiveWindowForWorkbook(workbookFullName);
            if (activeWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook window from ActiveWindow. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(activeWindow));
                return activeWindow;
            }

            try
            {
                workbook.Activate();
            }
            catch (Exception ex)
            {
                _logger.Debug(
                    nameof(ExcelWindowRecoveryService),
                    "ResolveWindow workbook.Activate failed. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", message="
                    + ex.Message);
            }

            LogWindowResolutionSnapshot("after-activate", workbook, reason, workbookFullName, allowWindowCreation);

            visibleWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
            if (visibleWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook visible window after workbook.Activate. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(visibleWindow));
                return visibleWindow;
            }

            workbookWindow = TryGetFirstWorkbookWindow(workbook);
            if (workbookWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook window from Workbook.Windows after workbook.Activate. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(workbookWindow));
                return workbookWindow;
            }

            applicationWindow = TryFindApplicationWindowForWorkbook(workbook, workbookFullName, visibleOnly: false);
            if (applicationWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook window from Application.Windows after workbook.Activate. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(applicationWindow));
                return applicationWindow;
            }

            activeWindow = TryGetActiveWindowForWorkbook(workbookFullName);
            if (activeWindow != null)
            {
                _logger.Info(
                    "ResolveWindow resolved workbook window from ActiveWindow after workbook.Activate. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", window="
                    + DescribeWindow(activeWindow));
                return activeWindow;
            }

            if (!allowWindowCreation)
            {
                _logger.Warn(
                    "ResolveWindow could not find an existing workbook window and skipped NewWindow. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return null;
            }

            return TryRecreateWindow(workbook, reason, workbookFullName);
        }

        private Excel.Window TryRecreateWindow(Excel.Workbook workbook, string reason, string workbookFullName)
        {
            try
            {
                workbook.Activate();
                if (workbook.Windows.Count > 0)
                {
                    _logger.Info(
                        "Workbook window recreated by activation. reason="
                        + (reason ?? string.Empty)
                        + ", workbook="
                        + (workbookFullName ?? string.Empty));
                    return workbook.Windows[1];
                }
            }
            catch (Exception ex)
            {
                _logger.Debug(
                    nameof(ExcelWindowRecoveryService),
                    "Workbook activation did not recreate a window. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", message="
                    + ex.Message);
            }

            try
            {
                Excel.Window createdWindow = workbook.NewWindow();
                if (createdWindow != null)
                {
                    _logger.Info(
                        "Workbook window recreated by NewWindow. reason="
                        + (reason ?? string.Empty)
                        + ", workbook="
                        + (workbookFullName ?? string.Empty)
                        + ", window="
                        + DescribeWindow(createdWindow));
                    return createdWindow;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "TryRecreateWindow failed. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty),
                    ex);
            }

            return null;
        }

        private Excel.Window TryGetFirstWorkbookWindow(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            try
            {
                int count = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= count; index++)
                {
                    Excel.Window window = workbook.Windows[index];
                    if (window != null)
                    {
                        return window;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("TryGetFirstWorkbookWindow failed.", ex);
            }

            return null;
        }

        private Excel.Window TryGetActiveWindowForWorkbook(string workbookFullName)
        {
            if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                return null;
            }

            try
            {
                Excel.Workbook activeWorkbook = _excelInteropService.GetActiveWorkbook();
                string activeWorkbookFullName = _excelInteropService.GetWorkbookFullName(activeWorkbook);
                if (!string.Equals(activeWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                return _excelInteropService.GetActiveWindow();
            }
            catch (Exception ex)
            {
                _logger.Error("TryGetActiveWindowForWorkbook failed.", ex);
                return null;
            }
        }

        private Excel.Window TryFindApplicationWindowForWorkbook(Excel.Workbook workbook, string workbookFullName, bool visibleOnly)
        {
            object windowsObject = null;
            dynamic windowsCollection = null;
            try
            {
                dynamic lateBoundApplication = _application;
                windowsCollection = lateBoundApplication.Windows;
                windowsObject = windowsCollection;
                if (windowsCollection == null)
                {
                    return null;
                }

                int count = Convert.ToInt32(windowsCollection.Count, CultureInfo.InvariantCulture);
                for (int index = 1; index <= count; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = windowsCollection[index] as Excel.Window;
                    }
                    catch
                    {
                        window = null;
                    }

                    if (window == null)
                    {
                        continue;
                    }

                    if (visibleOnly && !SafeWindowVisibleValue(window))
                    {
                        continue;
                    }

                    if (DoesWindowBelongToWorkbook(window, workbook, workbookFullName))
                    {
                        return window;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Debug(
                    nameof(ExcelWindowRecoveryService),
                    "TryFindApplicationWindowForWorkbook failed. workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", message="
                    + ex.Message);
            }
            finally
            {
                ReleaseComObject(windowsObject);
            }

            return null;
        }

        private bool DoesWindowBelongToWorkbook(Excel.Window window, Excel.Workbook workbook, string workbookFullName)
        {
            if (window == null)
            {
                return false;
            }

            object parent = null;
            try
            {
                dynamic lateBoundWindow = window;
                parent = lateBoundWindow.Parent;
                Excel.Workbook parentWorkbook = parent as Excel.Workbook;
                if (ReferenceEquals(parentWorkbook, workbook))
                {
                    return true;
                }

                string parentWorkbookFullName = _excelInteropService.GetWorkbookFullName(parentWorkbook);
                if (!string.IsNullOrWhiteSpace(parentWorkbookFullName)
                    && string.Equals(parentWorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch
            {
            }
            finally
            {
                if (!ReferenceEquals(parent, workbook))
                {
                    ReleaseComObject(parent);
                }
            }

            return CaptionMatchesWorkbookName(SafeWindowCaption(window), SafeWorkbookName(workbook));
        }

        private void LogWindowResolutionSnapshot(string stage, Excel.Workbook workbook, string reason, string workbookFullName, bool allowWindowCreation)
        {
            if (allowWindowCreation)
            {
                return;
            }

            _logger.Info(
                "ResolveWindow snapshot. stage="
                + (stage ?? string.Empty)
                + ", reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + (workbookFullName ?? string.Empty)
                + ", applicationWorkbooks.Count="
                + GetApplicationWorkbookCount().ToString(CultureInfo.InvariantCulture)
                + ", applicationWindows.Count="
                + GetApplicationWindowCount().ToString(CultureInfo.InvariantCulture)
                + ", workbook.Windows.Count="
                + GetWorkbookWindowCount(workbook).ToString(CultureInfo.InvariantCulture)
                + ", activeWorkbook="
                + SafeWorkbookFullName(_excelInteropService.GetActiveWorkbook())
                + ", activeWindow="
                + DescribeWindow(_excelInteropService.GetActiveWindow())
                + ", workbookWindows="
                + DescribeWorkbookWindows(workbook)
                + ", applicationWindows="
                + DescribeApplicationWindows(workbook, workbookFullName));
        }

        private int GetApplicationWorkbookCount()
        {
            try
            {
                return _application == null || _application.Workbooks == null ? 0 : _application.Workbooks.Count;
            }
            catch
            {
                return -1;
            }
        }

        private int GetApplicationWindowCount()
        {
            object windowsObject = null;
            dynamic windowsCollection = null;
            try
            {
                dynamic lateBoundApplication = _application;
                windowsCollection = lateBoundApplication.Windows;
                windowsObject = windowsCollection;
                return windowsCollection == null
                    ? 0
                    : Convert.ToInt32(windowsCollection.Count, CultureInfo.InvariantCulture);
            }
            catch
            {
                return -1;
            }
            finally
            {
                ReleaseComObject(windowsObject);
            }
        }

        private static int GetWorkbookWindowCount(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null || workbook.Windows == null ? 0 : workbook.Windows.Count;
            }
            catch
            {
                return -1;
            }
        }

        private string DescribeWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return "none";
            }

            try
            {
                int count = workbook.Windows == null ? 0 : workbook.Windows.Count;
                if (count == 0)
                {
                    return "none";
                }

                string[] descriptors = new string[count];
                for (int index = 1; index <= count; index++)
                {
                    descriptors[index - 1] = "index="
                        + index.ToString(CultureInfo.InvariantCulture)
                        + ",window="
                        + DescribeWindow(workbook.Windows[index]);
                }

                return string.Join(" | ", descriptors);
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }
        }

        private string DescribeApplicationWindows(Excel.Workbook workbook, string workbookFullName)
        {
            object windowsObject = null;
            dynamic windowsCollection = null;
            try
            {
                dynamic lateBoundApplication = _application;
                windowsCollection = lateBoundApplication.Windows;
                windowsObject = windowsCollection;
                if (windowsCollection == null)
                {
                    return "none";
                }

                int count = Convert.ToInt32(windowsCollection.Count, CultureInfo.InvariantCulture);
                if (count == 0)
                {
                    return "none";
                }

                string[] descriptors = new string[count];
                for (int index = 1; index <= count; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = windowsCollection[index] as Excel.Window;
                    }
                    catch
                    {
                        window = null;
                    }

                    descriptors[index - 1] = "index="
                        + index.ToString(CultureInfo.InvariantCulture)
                        + ",window="
                        + DescribeWindow(window)
                        + ",belongsToTarget="
                        + DoesWindowBelongToWorkbook(window, workbook, workbookFullName).ToString();
                }

                return string.Join(" | ", descriptors);
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }
            finally
            {
                ReleaseComObject(windowsObject);
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

        private static string DescribeWindow(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\",visible=\""
                + SafeWindowVisible(window)
                + "\",state=\""
                + SafeWindowState(window)
                + "\"";
        }

        private static string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWorkbookName(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null ? string.Empty : workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowCaption(Excel.Window window)
        {
            try
            {
                if (window == null)
                {
                    return string.Empty;
                }

                dynamic lateBoundWindow = window;
                return Convert.ToString(lateBoundWindow.Caption) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowVisible(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.Visible.ToString();
            }
            catch
            {
                return "error";
            }
        }

        private static bool SafeWindowVisibleValue(Excel.Window window)
        {
            try
            {
                return window != null && window.Visible;
            }
            catch
            {
                return false;
            }
        }

        private static string SafeWindowState(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.WindowState.ToString();
            }
            catch
            {
                return "error";
            }
        }

        private static bool CaptionMatchesWorkbookName(string caption, string workbookName)
        {
            if (string.IsNullOrWhiteSpace(caption) || string.IsNullOrWhiteSpace(workbookName))
            {
                return false;
            }

            if (string.Equals(caption, workbookName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return caption.StartsWith(workbookName + ":", StringComparison.OrdinalIgnoreCase);
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
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
