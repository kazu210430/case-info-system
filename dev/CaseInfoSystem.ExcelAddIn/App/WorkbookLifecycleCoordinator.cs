using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookLifecycleCoordinator
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Logger _logger;
        private readonly ExcelInteropService _excelInteropService;
        private readonly KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;
        private readonly AccountingWorkbookLifecycleService _accountingWorkbookLifecycleService;
        private readonly AccountingSheetControlService _accountingSheetControlService;
        private readonly WorkbookClipboardPreservationService _workbookClipboardPreservationService;
        private readonly TaskPaneManager _taskPaneManager;
        private readonly KernelHomeCoordinator _kernelHomeCoordinator;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Action<string, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly ICasePaneHostBridge _casePaneHostBridge;

        internal WorkbookLifecycleCoordinator(
            Logger logger,
            ExcelInteropService excelInteropService,
            KernelWorkbookLifecycleService kernelWorkbookLifecycleService,
            CaseWorkbookLifecycleService caseWorkbookLifecycleService,
            AccountingWorkbookLifecycleService accountingWorkbookLifecycleService,
            AccountingSheetControlService accountingSheetControlService,
            WorkbookClipboardPreservationService workbookClipboardPreservationService,
            TaskPaneManager taskPaneManager,
            KernelHomeCoordinator kernelHomeCoordinator,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPane,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            ICasePaneHostBridge casePaneHostBridge)
        {
            _logger = logger;
            _excelInteropService = excelInteropService;
            _kernelWorkbookLifecycleService = kernelWorkbookLifecycleService;
            _caseWorkbookLifecycleService = caseWorkbookLifecycleService;
            _accountingWorkbookLifecycleService = accountingWorkbookLifecycleService;
            _accountingSheetControlService = accountingSheetControlService;
            _workbookClipboardPreservationService = workbookClipboardPreservationService;
            _taskPaneManager = taskPaneManager;
            _kernelHomeCoordinator = kernelHomeCoordinator;
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _refreshTaskPane = refreshTaskPane;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException(nameof(casePaneHostBridge));
        }

        internal void OnWorkbookOpen(Excel.Workbook workbook)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ExcelEventBoundary action=fire event=WorkbookOpen workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "Excel WorkbookOpen fired. workbook="
                + SafeWorkbookFullName(workbook)
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WorkbookOpen workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "TaskPane event entry. event=WorkbookOpen, workbook="
                + SafeWorkbookFullName(workbook)
                + ", activeWorkbook="
                + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook())
                + ", activeWindowHwnd="
                + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
            NewCaseVisibilityObservation.Log(_logger, _excelInteropService, null, workbook, null, "WorkbookOpen-event", "WorkbookLifecycleCoordinator.OnWorkbookOpen");

            _handleExternalWorkbookDetected?.Invoke(workbook, "WorkbookOpen");
            _kernelWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _accountingWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook, AccountingInitialSheetSyncPolicy.WorkbookOpenEventName);
            _accountingSheetControlService?.EnsureVstoManagedControls(workbook);
            _caseWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _kernelHomeCoordinator?.HandleKernelWorkbookBecameAvailable("WorkbookOpen", workbook);
        }

        internal void OnWorkbookActivate(Excel.Workbook workbook)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ExcelEventBoundary action=fire event=WorkbookActivate workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "Excel WorkbookActivate fired. workbook="
                + SafeWorkbookFullName(workbook)
                + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WorkbookActivate workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "TaskPane event entry. event=WorkbookActivate, workbook="
                + SafeWorkbookFullName(workbook)
                + ", activeWorkbook="
                + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook())
                + ", activeWindowHwnd="
                + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
            NewCaseVisibilityObservation.Log(_logger, _excelInteropService, null, workbook, null, "WorkbookActivate-event", "WorkbookLifecycleCoordinator.OnWorkbookActivate");

            if (_casePaneHostBridge.ShouldIgnoreWorkbookActivateDuringCaseProtection(workbook))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=WorkbookEventCoordinator action=ignore-reentrant-activate event=WorkbookActivate workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveExcelState());
                return;
            }

            _handleExternalWorkbookDetected?.Invoke(workbook, "WorkbookActivate");
            _kernelWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _accountingWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook, AccountingInitialSheetSyncPolicy.WorkbookActivateEventName);
            _accountingSheetControlService?.EnsureVstoManagedControls(workbook);
            _caseWorkbookLifecycleService?.HandleWorkbookOpenedOrActivated(workbook);
            _kernelHomeCoordinator?.HandleKernelWorkbookBecameAvailable("WorkbookActivate", workbook);
            if (_shouldSuppressCasePaneRefresh != null && _shouldSuppressCasePaneRefresh("WorkbookActivate", workbook))
            {
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=WorkbookEventCoordinator action=suppress-refresh event=WorkbookActivate workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + FormatActiveExcelState());
                return;
            }

            _refreshTaskPane?.Invoke("WorkbookActivate", workbook, null);
        }

        internal void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            _logger?.Info(
                "Excel WorkbookBeforeClose fired. workbook="
                + SafeWorkbookFullName(workbook)
                + ", cancel="
                + cancel.ToString());

            if (cancel)
            {
                return;
            }

            _caseWorkbookLifecycleService?.HandleWorkbookBeforeClose(workbook, ref cancel);
            if (cancel)
            {
                return;
            }

            _kernelWorkbookLifecycleService?.HandleWorkbookBeforeClose(workbook, ref cancel);
            if (cancel)
            {
                return;
            }

            _workbookClipboardPreservationService?.PreserveCopiedValuesForClosingWorkbook(workbook);
            _accountingWorkbookLifecycleService?.HandleWorkbookBeforeClose(workbook);

            if (_accountingSheetControlService != null)
            {
                _accountingSheetControlService.RemoveWorkbookState(workbook);
            }

            if (_accountingWorkbookLifecycleService != null)
            {
                _accountingWorkbookLifecycleService.RemoveWorkbookState(workbook);
            }

            if (_caseWorkbookLifecycleService != null)
            {
                _caseWorkbookLifecycleService.RemoveWorkbookState(workbook);
            }

            if (_taskPaneManager != null)
            {
                _taskPaneManager.RemoveWorkbookPanes(workbook);
            }
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
        }

        private string FormatActiveExcelState()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            Excel.Window activeWindow = _excelInteropService == null ? null : _excelInteropService.GetActiveWindow();
            return "activeWorkbook=" + FormatWorkbookDescriptor(activeWorkbook) + ",activeWindow=" + FormatWindowDescriptor(activeWindow);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private string SafeWorkbookName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookName(workbook);
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\"";
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
                return Convert.ToString(lateBoundWindow.Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
