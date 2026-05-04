using System;
using System.Diagnostics;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookWindowVisibilityService
    {
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        internal WorkbookWindowVisibilityService(ExcelInteropService excelInteropService, Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal WorkbookWindowVisibilityEnsureResult EnsureVisible(Excel.Workbook workbook, string reason)
        {
            if (workbook == null)
            {
                return WorkbookWindowVisibilityEnsureResult.Create(
                    WorkbookWindowVisibilityEnsureOutcome.WorkbookMissing,
                    string.Empty,
                    null,
                    0L,
                    visibleAfterSet: null);
            }

            string workbookFullName = _excelInteropService.GetWorkbookFullName(workbook);
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                Excel.Window workbookWindow = _excelInteropService.GetFirstVisibleWindow(workbook);
                if (workbookWindow == null && workbook.Windows.Count > 0)
                {
                    workbookWindow = workbook.Windows[1];
                }

                if (workbookWindow == null)
                {
                    _logger.Warn(
                        "Workbook window visibility ensure skipped because workbook window could not be resolved. reason="
                        + (reason ?? string.Empty)
                        + ", workbook="
                        + workbookFullName
                        + ", elapsedMs="
                        + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                    return WorkbookWindowVisibilityEnsureResult.Create(
                        WorkbookWindowVisibilityEnsureOutcome.WindowUnresolved,
                        workbookFullName,
                        null,
                        stopwatch.ElapsedMilliseconds,
                        visibleAfterSet: null);
                }

                bool isVisible;
                try
                {
                    isVisible = workbookWindow.Visible;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        "Workbook window visibility ensure failed while reading Window.Visible. reason="
                        + (reason ?? string.Empty)
                        + ", workbook="
                        + workbookFullName
                        + ", elapsedMs="
                        + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture),
                        ex);
                    return WorkbookWindowVisibilityEnsureResult.Create(
                        WorkbookWindowVisibilityEnsureOutcome.VisibilityReadFailed,
                        workbookFullName,
                        workbookWindow,
                        stopwatch.ElapsedMilliseconds,
                        visibleAfterSet: null);
                }

                if (isVisible)
                {
                    return WorkbookWindowVisibilityEnsureResult.Create(
                        WorkbookWindowVisibilityEnsureOutcome.AlreadyVisible,
                        workbookFullName,
                        workbookWindow,
                        stopwatch.ElapsedMilliseconds,
                        visibleAfterSet: true);
                }

                workbookWindow.Visible = true;
                bool? visibleAfterSet = null;
                try
                {
                    visibleAfterSet = workbookWindow.Visible;
                }
                catch
                {
                    visibleAfterSet = null;
                }

                return WorkbookWindowVisibilityEnsureResult.Create(
                    WorkbookWindowVisibilityEnsureOutcome.MadeVisible,
                    workbookFullName,
                    workbookWindow,
                    stopwatch.ElapsedMilliseconds,
                    visibleAfterSet);
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Workbook window visibility ensure failed. reason="
                    + (reason ?? string.Empty)
                    + ", workbook="
                    + workbookFullName
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture),
                    ex);
                return WorkbookWindowVisibilityEnsureResult.Create(
                    WorkbookWindowVisibilityEnsureOutcome.Failed,
                    workbookFullName,
                    null,
                    stopwatch.ElapsedMilliseconds,
                    visibleAfterSet: null);
            }
        }
    }

    internal enum WorkbookWindowVisibilityEnsureOutcome
    {
        WorkbookMissing,
        WindowUnresolved,
        VisibilityReadFailed,
        AlreadyVisible,
        MadeVisible,
        Failed,
    }

    internal sealed class WorkbookWindowVisibilityEnsureResult
    {
        private WorkbookWindowVisibilityEnsureResult(
            WorkbookWindowVisibilityEnsureOutcome outcome,
            string workbookFullName,
            Excel.Window window,
            long elapsedMilliseconds,
            bool? visibleAfterSet)
        {
            Outcome = outcome;
            WorkbookFullName = workbookFullName ?? string.Empty;
            Window = window;
            WindowHwnd = SafeWindowHwnd(window);
            ElapsedMilliseconds = elapsedMilliseconds;
            VisibleAfterSet = visibleAfterSet;
        }

        internal WorkbookWindowVisibilityEnsureOutcome Outcome { get; }

        internal string WorkbookFullName { get; }

        internal Excel.Window Window { get; }

        internal string WindowHwnd { get; }

        internal long ElapsedMilliseconds { get; }

        internal bool? VisibleAfterSet { get; }

        internal static WorkbookWindowVisibilityEnsureResult Create(
            WorkbookWindowVisibilityEnsureOutcome outcome,
            string workbookFullName,
            Excel.Window window,
            long elapsedMilliseconds,
            bool? visibleAfterSet)
        {
            return new WorkbookWindowVisibilityEnsureResult(
                outcome,
                workbookFullName,
                window,
                elapsedMilliseconds,
                visibleAfterSet);
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null
                    ? string.Empty
                    : Convert.ToString(window.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
