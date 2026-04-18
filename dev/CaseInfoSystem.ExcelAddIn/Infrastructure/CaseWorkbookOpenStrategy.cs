using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CaseWorkbookOpenStrategy
    {
        private readonly Excel.Application _application;
        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly Logger _logger;

        internal CaseWorkbookOpenStrategy(Excel.Application application, WorkbookRoleResolver workbookRoleResolver, Logger logger)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException(nameof(workbookRoleResolver));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal void RegisterKnownCasePath(string caseWorkbookPath)
        {
            _workbookRoleResolver.RegisterKnownCasePath(caseWorkbookPath);
        }

        internal Excel.Workbook OpenVisibleWorkbook(string caseWorkbookPath)
        {
            _logger.Info("Case workbook open visible requested. path=" + (caseWorkbookPath ?? string.Empty));
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Window previousActiveWindow = null;
            try
            {
                previousActiveWindow = _application.ActiveWindow;
                Excel.Workbook workbook = _application.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
                _logger.Info("Case workbook visible open completed. path=" + (caseWorkbookPath ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());
                _workbookRoleResolver.RegisterKnownCaseWorkbook(workbook);
                HideOpenedWorkbookWindow(workbook);
                RestorePreviousWindow(previousActiveWindow);
                _logger.Info("Case workbook visible open post-processing completed. path=" + (caseWorkbookPath ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());
                return workbook;
            }
            catch
            {
                RestorePreviousWindow(previousActiveWindow);
                throw;
            }
        }

        internal HiddenCaseWorkbookSession OpenHiddenWorkbook(string caseWorkbookPath)
        {
            _logger.Info("Case workbook open hidden requested. path=" + (caseWorkbookPath ?? string.Empty));
            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Application hiddenApplication = new Excel.Application();
            _logger.Info("Case workbook hidden Excel application created. path=" + (caseWorkbookPath ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());
            hiddenApplication.Visible = false;
            hiddenApplication.DisplayAlerts = false;
            hiddenApplication.ScreenUpdating = false;
            hiddenApplication.EnableEvents = false;
            _logger.Info("Case workbook hidden Excel application configured. path=" + (caseWorkbookPath ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());

            Excel.Workbook workbook = hiddenApplication.Workbooks.Open(caseWorkbookPath, ReadOnly: false, UpdateLinks: 0);
            _logger.Info(
                "Case workbook hidden Excel session opened. path="
                + (caseWorkbookPath ?? string.Empty)
                + ", appHwnd="
                + SafeApplicationHwnd(hiddenApplication)
                + ", elapsedMs="
                + stopwatch.ElapsedMilliseconds.ToString());
            return new HiddenCaseWorkbookSession(hiddenApplication, workbook);
        }

        private static string SafeApplicationHwnd(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : Convert.ToString(application.Hwnd) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void HideOpenedWorkbookWindow(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null)
                    {
                        window.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("HideOpenedWorkbookWindow failed.", ex);
            }
        }

        private void RestorePreviousWindow(Excel.Window previousActiveWindow)
        {
            if (previousActiveWindow == null)
            {
                return;
            }

            try
            {
                previousActiveWindow.Visible = true;
                previousActiveWindow.Activate();
            }
            catch (Exception ex)
            {
                _logger.Error("RestorePreviousWindow failed.", ex);
            }
        }

        internal sealed class HiddenCaseWorkbookSession
        {
            internal HiddenCaseWorkbookSession(Excel.Application application, Excel.Workbook workbook)
            {
                Application = application ?? throw new ArgumentNullException(nameof(application));
                Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            }

            internal Excel.Application Application { get; }

            internal Excel.Workbook Workbook { get; }
        }
    }
}
