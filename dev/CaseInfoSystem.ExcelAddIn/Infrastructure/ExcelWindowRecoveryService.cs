using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class ExcelWindowRecoveryService
    {
        private static readonly IntPtr HwndTopMost = new IntPtr(-1);
        private static readonly IntPtr HwndNoTopMost = new IntPtr(-2);
        private const int SwHide = 0;
        private const int SwShow = 5;
        private const int SwRestore = 9;
        private const uint SwpNoMove = 0x0002;
        private const uint SwpNoSize = 0x0001;
        private const uint SwpShowWindow = 0x0040;
        private const string NoActiveWorkbook = "none";
        private const string NoActiveWindow = "none";
        private const string UnresolvedWindow = "unresolved";
        private const string ReadFailed = "read-failed";

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private sealed class WindowMutationTraceContext
        {
            internal WindowMutationTraceContext(string workbookFullName, string reason)
            {
                WorkbookFullName = workbookFullName ?? string.Empty;
                Reason = reason ?? string.Empty;
            }

            internal string WorkbookFullName { get; }

            internal string Reason { get; }

            internal WindowMutationSnapshot PreviousSnapshot { get; set; }
        }

        private sealed class WindowMutationSnapshot
        {
            internal string AppHwnd { get; set; }

            internal string WindowHwnd { get; set; }

            internal string ActiveWorkbookFullName { get; set; }

            internal string ActiveWindowHwnd { get; set; }

            internal string Visible { get; set; }

            internal string WindowState { get; set; }

            internal string Left { get; set; }

            internal string Top { get; set; }

            internal string Width { get; set; }

            internal string Height { get; set; }

            internal string Failure { get; set; }
        }

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
            WindowMutationTraceContext traceContext = new WindowMutationTraceContext(workbookFullName, reason);
            bool recoveredScreenUpdating = EnsureScreenUpdatingEnabled(reason, workbookFullName);
            Excel.Window window = ResolveWindow(workbook, traceContext);
            TraceWindowMutation(traceContext, "recovery-start", window);
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
                && EnsureWindowVisible(window, traceContext);
            bool recoveredWindowState = EnsureWindowRestored(window, traceContext);
            bool recoveredApplicationVisibility = EnsureApplicationVisible(window, traceContext);

            if (activateWindow)
            {
                TraceWindowMutation(traceContext, "window-activate-before", window);
                try
                {
                    window.Activate();
                    TraceWindowMutation(traceContext, "window-activate-after", window);
                }
                catch (Exception ex)
                {
                    TraceWindowMutation(traceContext, "window-activate-failed", window);
                    _logger.Debug(nameof(ExcelWindowRecoveryService), "window.Activate failed but recovery continues. reason=" + (reason ?? string.Empty) + ", workbook=" + workbookFullName + ", message=" + ex.Message);
                }
            }

            if (bringToFront
                && (ensureWindowVisible || window.Visible))
            {
                PromoteExcelWindow(window, traceContext);
            }

            TraceWindowMutation(traceContext, "recovery-complete", window);
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
        internal bool HideApplicationWindow(string reason, string workbookFullName)
        {
            try
            {
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                ShowWindow(hwnd, SwHide);
                _application.Visible = false;
                _logger.Info(
                    "Excel application window hidden. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("HideApplicationWindow failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        internal bool ShowApplicationWindow(string reason, string workbookFullName)
        {
            WindowMutationTraceContext traceContext = new WindowMutationTraceContext(workbookFullName, reason);
            try
            {
                TraceWindowMutation(traceContext, "show-application-visible-before", GetCurrentActiveWindowForTracing());
                _application.Visible = true;
                TraceWindowMutation(traceContext, "show-application-visible-after", GetCurrentActiveWindowForTracing());

                IntPtr hwnd = new IntPtr(_application.Hwnd);
                TraceWindowMutation(traceContext, "show-application-showwindow-restore-before", GetCurrentActiveWindowForTracing());
                ShowWindow(hwnd, SwRestore);
                TraceWindowMutation(traceContext, "show-application-showwindow-restore-after", GetCurrentActiveWindowForTracing());

                ShowWindow(hwnd, SwShow);
                TraceWindowMutation(traceContext, "show-application-showwindow-after", GetCurrentActiveWindowForTracing());

                _logger.Info(
                    "Excel application window shown. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty));
                return true;
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "show-application-failed", GetCurrentActiveWindowForTracing());
                _logger.Error("ShowApplicationWindow failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        internal bool TryBringApplicationToForeground(string reason, string workbookFullName)
        {
            WindowMutationTraceContext traceContext = new WindowMutationTraceContext(workbookFullName, reason);
            try
            {
                TraceWindowMutation(traceContext, "application-foreground-before", GetCurrentActiveWindowForTracing());
                IntPtr hwnd = new IntPtr(_application.Hwnd);
                bool promoted = SetForegroundWindow(hwnd);
                TraceWindowMutation(traceContext, "application-foreground-after", GetCurrentActiveWindowForTracing());
                _logger.Info(
                    "Excel application foreground requested. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + (workbookFullName ?? string.Empty)
                    + ", promoted="
                    + promoted.ToString());
                return promoted;
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "application-foreground-failed", GetCurrentActiveWindowForTracing());
                _logger.Error("TryBringApplicationToForeground failed. reason=" + (reason ?? string.Empty) + ", workbook=" + (workbookFullName ?? string.Empty), ex);
                return false;
            }
        }

        /// <summary>
        internal bool EnsureApplicationVisible(string reason, string workbookFullName)
        {
            WindowMutationTraceContext traceContext = new WindowMutationTraceContext(workbookFullName, reason);
            return EnsureApplicationVisible(GetCurrentActiveWindowForTracing(), traceContext);
        }

        private bool EnsureApplicationVisible(Excel.Window window, WindowMutationTraceContext traceContext)
        {
            try
            {
                TraceWindowMutation(traceContext, "application-visible-before", window);
                if (_application.Visible)
                {
                    TraceWindowMutation(traceContext, "application-visible-skip", window);
                    return false;
                }

                _application.Visible = true;
                TraceWindowMutation(traceContext, "application-visible-after", window);

                IntPtr hwnd = new IntPtr(_application.Hwnd);
                TraceWindowMutation(traceContext, "application-showwindow-restore-before", window);
                ShowWindow(hwnd, SwRestore);
                TraceWindowMutation(traceContext, "application-showwindow-restore-after", window);

                ShowWindow(hwnd, SwShow);
                TraceWindowMutation(traceContext, "application-showwindow-after", window);

                _logger.Info(
                    "Excel application visibility recovered. reason="
                    + traceContext.Reason
                    + ", workbook="
                    + traceContext.WorkbookFullName);
                return true;
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "application-visible-failed", window);
                _logger.Error("EnsureApplicationVisible failed. reason=" + traceContext.Reason + ", workbook=" + traceContext.WorkbookFullName, ex);
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
        private Excel.Window ResolveWindow(Excel.Workbook workbook, WindowMutationTraceContext traceContext)
        {
            Excel.Window visibleWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
            if (visibleWindow != null)
            {
                return visibleWindow;
            }

            try
            {
                if (workbook.Windows.Count > 0)
                {
                    return workbook.Windows[1];
                }
            }
            catch (Exception ex)
            {
                _logger.Error("ResolveWindow failed.", ex);
            }

            return TryRecreateWindow(workbook, traceContext);
        }

        private Excel.Window TryRecreateWindow(Excel.Workbook workbook, WindowMutationTraceContext traceContext)
        {
            try
            {
                TraceWindowMutation(traceContext, "workbook-activate-before", TryGetFirstWindowWithoutMutation(workbook));
                workbook.Activate();
                Excel.Window activatedWindow = TryGetFirstWindowWithoutMutation(workbook);
                TraceWindowMutation(traceContext, "workbook-activate-after", activatedWindow);
                if (activatedWindow != null)
                {
                    _logger.Info(
                        "Workbook window recreated by activation. reason="
                        + traceContext.Reason
                        + ", workbook="
                        + traceContext.WorkbookFullName);
                    return activatedWindow;
                }
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "workbook-activate-failed", TryGetFirstWindowWithoutMutation(workbook));
                _logger.Debug(
                    nameof(ExcelWindowRecoveryService),
                    "Workbook activation did not recreate a window. reason="
                    + traceContext.Reason
                    + ", workbook="
                    + traceContext.WorkbookFullName
                    + ", message="
                    + ex.Message);
            }

            try
            {
                TraceWindowMutation(traceContext, "newwindow-before", TryGetFirstWindowWithoutMutation(workbook));
                Excel.Window createdWindow = workbook.NewWindow();
                TraceWindowMutation(traceContext, "newwindow-after", createdWindow);
                if (createdWindow != null)
                {
                    _logger.Info(
                        "Workbook window recreated by NewWindow. reason="
                        + traceContext.Reason
                        + ", workbook="
                        + traceContext.WorkbookFullName);
                    return createdWindow;
                }
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "newwindow-failed", TryGetFirstWindowWithoutMutation(workbook));
                _logger.Error(
                    "TryRecreateWindow failed. reason="
                    + traceContext.Reason
                    + ", workbook="
                    + traceContext.WorkbookFullName,
                    ex);
            }

            return null;
        }

        /// <summary>
        private bool EnsureWindowVisible(Excel.Window window, WindowMutationTraceContext traceContext)
        {
            TraceWindowMutation(traceContext, "window-visible-before", window);
            try
            {
                if (window.Visible)
                {
                    TraceWindowMutation(traceContext, "window-visible-skip", window);
                    return false;
                }

                window.Visible = true;
                TraceWindowMutation(traceContext, "window-visible-after", window);
                _logger.Info(
                    "Workbook window visibility recovered. reason="
                    + traceContext.Reason
                    + ", workbook="
                    + traceContext.WorkbookFullName);
                return true;
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "window-visible-failed", window);
                _logger.Error("EnsureWindowVisible failed. reason=" + traceContext.Reason + ", workbook=" + traceContext.WorkbookFullName, ex);
                return false;
            }
        }

        private bool EnsureWindowRestored(Excel.Window window, WindowMutationTraceContext traceContext)
        {
            TraceWindowMutation(traceContext, "windowstate-xlNormal-before", window);
            try
            {
                if (window.WindowState != Excel.XlWindowState.xlMinimized)
                {
                    TraceWindowMutation(traceContext, "windowstate-xlNormal-skip", window);
                    return false;
                }

                window.WindowState = Excel.XlWindowState.xlNormal;
                TraceWindowMutation(traceContext, "windowstate-xlNormal-after", window);
                _logger.Info(
                    "Workbook window state recovered. reason="
                    + traceContext.Reason
                    + ", workbook="
                    + traceContext.WorkbookFullName);
                return true;
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "windowstate-xlNormal-failed", window);
                _logger.Error("EnsureWindowRestored failed. reason=" + traceContext.Reason + ", workbook=" + traceContext.WorkbookFullName, ex);
                return false;
            }
        }

        /// <summary>
        private void PromoteExcelWindow(Excel.Window window, WindowMutationTraceContext traceContext)
        {
            try
            {
                TraceWindowMutation(traceContext, "promote-start", window);
                PromoteWindow(new IntPtr(window.Hwnd), window, traceContext);
                TraceWindowMutation(traceContext, "promote-after", window);
                _logger.Info(
                    "Excel window promoted after recovery. reason="
                    + traceContext.Reason
                    + ", workbook="
                    + traceContext.WorkbookFullName);
            }
            catch (Exception ex)
            {
                TraceWindowMutation(traceContext, "promote-failed", window);
                _logger.Debug(nameof(ExcelWindowRecoveryService), "PromoteWindow failed but recovery continues. reason=" + traceContext.Reason + ", workbook=" + traceContext.WorkbookFullName + ", message=" + ex.Message);
            }
        }

        /// <summary>
        private void PromoteWindow(IntPtr hwnd, Excel.Window window, WindowMutationTraceContext traceContext)
        {
            if (hwnd == IntPtr.Zero)
            {
                TraceWindowMutation(traceContext, "promote-skip-zero-hwnd", window);
                return;
            }

            TraceWindowMutation(traceContext, "promote-showwindow-restore-before", window);
            ShowWindow(hwnd, SwRestore);
            TraceWindowMutation(traceContext, "promote-showwindow-restore-after", window);
            SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpShowWindow);
            SetWindowPos(hwnd, HwndNoTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpShowWindow);
            TraceWindowMutation(traceContext, "promote-setforeground-before", window);
            SetForegroundWindow(hwnd);
            TraceWindowMutation(traceContext, "promote-setforeground-after", window);
        }

        private void TraceWindowMutation(WindowMutationTraceContext traceContext, string step, Excel.Window window)
        {
            if (traceContext == null)
            {
                return;
            }

            WindowMutationSnapshot snapshot = CaptureWindowMutationSnapshot(window);
            string changedFields = DescribeChangedFields(traceContext.PreviousSnapshot, snapshot);
            StringBuilder builder = new StringBuilder();
            builder.Append("Excel window recovery mutation trace. reason=")
                .Append(traceContext.Reason)
                .Append(", step=")
                .Append(step ?? string.Empty)
                .Append(", workbookFullName=")
                .Append(traceContext.WorkbookFullName)
                .Append(", appHwnd=")
                .Append(snapshot.AppHwnd)
                .Append(", windowHwnd=")
                .Append(snapshot.WindowHwnd)
                .Append(", activeWorkbookFullName=")
                .Append(snapshot.ActiveWorkbookFullName)
                .Append(", activeWindowHwnd=")
                .Append(snapshot.ActiveWindowHwnd)
                .Append(", visible=")
                .Append(snapshot.Visible)
                .Append(", windowState=")
                .Append(snapshot.WindowState)
                .Append(", left=")
                .Append(snapshot.Left)
                .Append(", top=")
                .Append(snapshot.Top)
                .Append(", width=")
                .Append(snapshot.Width)
                .Append(", height=")
                .Append(snapshot.Height)
                .Append(", changedFields=")
                .Append(changedFields)
                .Append(", failure=")
                .Append(snapshot.Failure);
            _logger.Info(builder.ToString());
            traceContext.PreviousSnapshot = snapshot;
        }

        private WindowMutationSnapshot CaptureWindowMutationSnapshot(Excel.Window window)
        {
            List<string> failures = new List<string>();
            WindowMutationSnapshot snapshot = new WindowMutationSnapshot();
            snapshot.AppHwnd = ReadApplicationHwnd(failures);

            IntPtr windowHwnd = IntPtr.Zero;
            snapshot.WindowHwnd = ReadWindowHwnd(window, failures, "windowHwnd", UnresolvedWindow, out windowHwnd);
            snapshot.ActiveWorkbookFullName = ReadActiveWorkbookFullName(failures);
            snapshot.ActiveWindowHwnd = ReadActiveWindowHwnd(failures);
            snapshot.Visible = ReadWindowVisible(window, failures);
            snapshot.WindowState = ReadWindowState(window, failures);
            ReadWindowRect(window, windowHwnd, snapshot, failures);
            snapshot.Failure = failures.Count == 0
                ? "none"
                : string.Join("|", failures.ToArray());
            return snapshot;
        }

        private string ReadApplicationHwnd(List<string> failures)
        {
            try
            {
                return FormatHwnd(new IntPtr(_application.Hwnd));
            }
            catch (Exception ex)
            {
                failures.Add("appHwnd:" + ex.GetType().Name);
                return ReadFailed;
            }
        }

        private string ReadActiveWorkbookFullName(List<string> failures)
        {
            try
            {
                Excel.Workbook activeWorkbook = _application.ActiveWorkbook;
                if (activeWorkbook == null)
                {
                    return NoActiveWorkbook;
                }

                return activeWorkbook.FullName ?? string.Empty;
            }
            catch (Exception ex)
            {
                failures.Add("activeWorkbookFullName:" + ex.GetType().Name);
                return ReadFailed;
            }
        }

        private string ReadActiveWindowHwnd(List<string> failures)
        {
            try
            {
                Excel.Window activeWindow = _application.ActiveWindow;
                if (activeWindow == null)
                {
                    return NoActiveWindow;
                }

                return FormatHwnd(new IntPtr(activeWindow.Hwnd));
            }
            catch (Exception ex)
            {
                failures.Add("activeWindowHwnd:" + ex.GetType().Name);
                return ReadFailed;
            }
        }

        private string ReadWindowHwnd(Excel.Window window, List<string> failures, string fieldName, string nullValue, out IntPtr hwnd)
        {
            hwnd = IntPtr.Zero;
            if (window == null)
            {
                return nullValue;
            }

            try
            {
                hwnd = new IntPtr(window.Hwnd);
                return FormatHwnd(hwnd);
            }
            catch (Exception ex)
            {
                failures.Add(fieldName + ":" + ex.GetType().Name);
                return ReadFailed;
            }
        }

        private string ReadWindowVisible(Excel.Window window, List<string> failures)
        {
            if (window == null)
            {
                return UnresolvedWindow;
            }

            try
            {
                return window.Visible.ToString();
            }
            catch (Exception ex)
            {
                failures.Add("visible:" + ex.GetType().Name);
                return ReadFailed;
            }
        }

        private string ReadWindowState(Excel.Window window, List<string> failures)
        {
            if (window == null)
            {
                return UnresolvedWindow;
            }

            try
            {
                return window.WindowState.ToString();
            }
            catch (Exception ex)
            {
                failures.Add("windowState:" + ex.GetType().Name);
                return ReadFailed;
            }
        }

        private void ReadWindowRect(Excel.Window window, IntPtr hwnd, WindowMutationSnapshot snapshot, List<string> failures)
        {
            if (window == null)
            {
                snapshot.Left = UnresolvedWindow;
                snapshot.Top = UnresolvedWindow;
                snapshot.Width = UnresolvedWindow;
                snapshot.Height = UnresolvedWindow;
                return;
            }

            if (hwnd != IntPtr.Zero && TryReadNativeWindowRect(hwnd, snapshot, failures))
            {
                return;
            }

            TryReadExcelWindowRect(window, snapshot, failures);
        }

        private bool TryReadNativeWindowRect(IntPtr hwnd, WindowMutationSnapshot snapshot, List<string> failures)
        {
            NativeRect rect;
            if (!GetWindowRect(hwnd, out rect))
            {
                failures.Add("windowRect:Win32Error" + Marshal.GetLastWin32Error().ToString(CultureInfo.InvariantCulture));
                return false;
            }

            snapshot.Left = rect.Left.ToString(CultureInfo.InvariantCulture);
            snapshot.Top = rect.Top.ToString(CultureInfo.InvariantCulture);
            snapshot.Width = (rect.Right - rect.Left).ToString(CultureInfo.InvariantCulture);
            snapshot.Height = (rect.Bottom - rect.Top).ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private void TryReadExcelWindowRect(Excel.Window window, WindowMutationSnapshot snapshot, List<string> failures)
        {
            try
            {
                snapshot.Left = FormatNumeric(window.Left);
                snapshot.Top = FormatNumeric(window.Top);
                snapshot.Width = FormatNumeric(window.Width);
                snapshot.Height = FormatNumeric(window.Height);
            }
            catch (Exception ex)
            {
                failures.Add("windowRectFallback:" + ex.GetType().Name);
                snapshot.Left = ReadFailed;
                snapshot.Top = ReadFailed;
                snapshot.Width = ReadFailed;
                snapshot.Height = ReadFailed;
            }
        }

        private static string DescribeChangedFields(WindowMutationSnapshot previousSnapshot, WindowMutationSnapshot currentSnapshot)
        {
            if (currentSnapshot == null)
            {
                return "none";
            }

            if (previousSnapshot == null)
            {
                return "initial";
            }

            List<string> changedFields = new List<string>();
            AddChangedField(changedFields, "appHwnd", previousSnapshot.AppHwnd, currentSnapshot.AppHwnd);
            AddChangedField(changedFields, "windowHwnd", previousSnapshot.WindowHwnd, currentSnapshot.WindowHwnd);
            AddChangedField(changedFields, "activeWorkbookFullName", previousSnapshot.ActiveWorkbookFullName, currentSnapshot.ActiveWorkbookFullName);
            AddChangedField(changedFields, "activeWindowHwnd", previousSnapshot.ActiveWindowHwnd, currentSnapshot.ActiveWindowHwnd);
            AddChangedField(changedFields, "visible", previousSnapshot.Visible, currentSnapshot.Visible);
            AddChangedField(changedFields, "windowState", previousSnapshot.WindowState, currentSnapshot.WindowState);
            AddChangedField(changedFields, "left", previousSnapshot.Left, currentSnapshot.Left);
            AddChangedField(changedFields, "top", previousSnapshot.Top, currentSnapshot.Top);
            AddChangedField(changedFields, "width", previousSnapshot.Width, currentSnapshot.Width);
            AddChangedField(changedFields, "height", previousSnapshot.Height, currentSnapshot.Height);
            return changedFields.Count == 0
                ? "none"
                : string.Join("|", changedFields.ToArray());
        }

        private static void AddChangedField(List<string> changedFields, string fieldName, string previousValue, string currentValue)
        {
            if (string.Equals(previousValue ?? string.Empty, currentValue ?? string.Empty, StringComparison.Ordinal))
            {
                return;
            }

            changedFields.Add(fieldName ?? string.Empty);
        }

        private Excel.Window GetCurrentActiveWindowForTracing()
        {
            try
            {
                return _application.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private Excel.Window TryGetFirstWindowWithoutMutation(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            try
            {
                foreach (Excel.Window candidate in workbook.Windows)
                {
                    if (candidate != null && candidate.Visible)
                    {
                        return candidate;
                    }
                }
            }
            catch
            {
            }

            try
            {
                return workbook.Windows.Count > 0
                    ? workbook.Windows[1]
                    : null;
            }
            catch
            {
                return null;
            }
        }

        private static string FormatHwnd(IntPtr hwnd)
        {
            return hwnd.ToInt64().ToString(CultureInfo.InvariantCulture);
        }

        private static string FormatNumeric(object value)
        {
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }
    }
}
