using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class VstoEventAdapter
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Excel.Application _application;
        private readonly Logger _logger;
        private readonly bool _subscribeSheetActivate;
        private readonly bool _subscribeSheetSelectionChange;
        private readonly bool _subscribeSheetChange;
        private ApplicationEventSubscriptionService _applicationEventSubscriptionService;
        private AddInStartupBoundaryCoordinator _startupBoundaryCoordinator;
        private ExcelInteropService _excelInteropService;
        private KernelWorkbookService _kernelWorkbookService;
        private KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private WorkbookLifecycleCoordinator _workbookLifecycleCoordinator;
        private WorkbookEventCoordinator _workbookEventCoordinator;
        private SheetEventCoordinator _sheetEventCoordinator;
        private WindowActivatePaneHandlingService _windowActivatePaneHandlingService;

        internal VstoEventAdapter(
            Excel.Application application,
            Logger logger,
            bool subscribeSheetActivate,
            bool subscribeSheetSelectionChange,
            bool subscribeSheetChange)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _logger = logger;
            _subscribeSheetActivate = subscribeSheetActivate;
            _subscribeSheetSelectionChange = subscribeSheetSelectionChange;
            _subscribeSheetChange = subscribeSheetChange;
        }

        internal void Configure(
            AddInStartupBoundaryCoordinator startupBoundaryCoordinator,
            ExcelInteropService excelInteropService,
            KernelWorkbookService kernelWorkbookService,
            KernelWorkbookLifecycleService kernelWorkbookLifecycleService,
            WorkbookLifecycleCoordinator workbookLifecycleCoordinator,
            WorkbookEventCoordinator workbookEventCoordinator,
            SheetEventCoordinator sheetEventCoordinator,
            WindowActivatePaneHandlingService windowActivatePaneHandlingService)
        {
            _startupBoundaryCoordinator = startupBoundaryCoordinator;
            _excelInteropService = excelInteropService;
            _kernelWorkbookService = kernelWorkbookService;
            _kernelWorkbookLifecycleService = kernelWorkbookLifecycleService;
            _workbookLifecycleCoordinator = workbookLifecycleCoordinator;
            _workbookEventCoordinator = workbookEventCoordinator;
            _sheetEventCoordinator = sheetEventCoordinator;
            _windowActivatePaneHandlingService = windowActivatePaneHandlingService;
            _applicationEventSubscriptionService = new ApplicationEventSubscriptionService(
                _application,
                Application_WorkbookOpen,
                Application_WorkbookActivate,
                Application_WorkbookBeforeSave,
                Application_WorkbookBeforeClose,
                Application_WindowActivate,
                Application_SheetActivate,
                Application_SheetSelectionChange,
                Application_SheetChange,
                Application_AfterCalculate,
                _subscribeSheetActivate,
                _subscribeSheetSelectionChange,
                _subscribeSheetChange);
        }

        internal void SubscribeApplicationEvents()
        {
            _applicationEventSubscriptionService?.Subscribe();
        }

        internal void UnsubscribeApplicationEvents()
        {
            _applicationEventSubscriptionService?.Unsubscribe();
        }

        internal string FormatActiveExcelState()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            Excel.Window activeWindow = _excelInteropService == null ? null : _excelInteropService.GetActiveWindow();
            return "activeWorkbook=" + FormatWorkbookDescriptor(activeWorkbook) + ",activeWindow=" + FormatWindowDescriptor(activeWindow);
        }

        internal void ClearKernelSheetCommandCell(Excel.Range commandCell)
        {
            if (commandCell == null)
            {
                return;
            }

            bool previousEnableEvents = _application.EnableEvents;
            try
            {
                _application.EnableEvents = false;
                commandCell.Value2 = string.Empty;
            }
            finally
            {
                _application.EnableEvents = previousEnableEvents;
            }
        }

        internal void HandleWindowActivateEvent(WindowActivateTaskPaneTriggerFacts triggerFacts)
        {
            if (triggerFacts == null)
            {
                triggerFacts = CaptureWindowActivateTaskPaneTriggerFacts(
                    null,
                    null,
                    "ThisAddIn.HandleWindowActivateEvent.NullFacts");
            }

            Excel.Workbook workbook = triggerFacts.Workbook;
            Excel.Window window = triggerFacts.Window;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=WorkbookEventCoordinator action=enter event=WindowActivate triggerRole=TaskPaneDisplayRefreshTrigger workbook="
                + triggerFacts.WorkbookDescriptor
                + ", eventWindow="
                + triggerFacts.WindowDescriptor
                + ", activeState="
                + triggerFacts.ActiveState
                + ", captureOwner="
                + triggerFacts.CaptureOwner);
            _logger?.Info("TaskPane event entry. event=WindowActivate, workbook=" + SafeWorkbookFullName(workbook) + ", windowHwnd=" + SafeWindowHwnd(window) + ", activeWorkbook=" + SafeWorkbookFullName(_excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook()) + ", activeWindowHwnd=" + SafeWindowHwnd(_excelInteropService == null ? null : _excelInteropService.GetActiveWindow()));
            NewCaseVisibilityObservation.Log(_logger, _excelInteropService, _application, workbook, window, "WindowActivate-event", "ThisAddIn.HandleWindowActivateEvent");
            _windowActivatePaneHandlingService?.Handle(triggerFacts);
        }

        internal void HandleWindowActivateEvent(Excel.Workbook workbook, Excel.Window window)
        {
            HandleWindowActivateEvent(CaptureWindowActivateTaskPaneTriggerFacts(
                workbook,
                window,
                "ThisAddIn.HandleWindowActivateEvent.Legacy"));
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            _startupBoundaryCoordinator?.MarkWorkbookOpenObserved();
            EnsureKernelFlickerTraceForWorkbookOpen(workbook);
            EventBoundaryGuard.Execute(_logger, nameof(Application_WorkbookOpen), () => _workbookLifecycleCoordinator?.OnWorkbookOpen(workbook));
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_WorkbookActivate), () => _workbookLifecycleCoordinator?.OnWorkbookActivate(workbook));
        }

        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_WindowActivate), () =>
            {
                WindowActivateTaskPaneTriggerFacts triggerFacts = CaptureWindowActivateTaskPaneTriggerFacts(
                    workbook,
                    window,
                    "ThisAddIn.Application_WindowActivate");
                _logger?.Info(
                    KernelFlickerTracePrefix
                    + " source=ExcelEventBoundary action=fire event=WindowActivate workbook="
                    + triggerFacts.WorkbookDescriptor
                    + ", eventWindow="
                    + triggerFacts.WindowDescriptor
                    + ", activeState="
                    + triggerFacts.ActiveState
                    + ", triggerRole=TaskPaneDisplayRefreshTrigger");
                _logger?.Info(
                    "Excel WindowActivate fired. workbook="
                    + triggerFacts.WorkbookFullName
                    + ", windowHwnd="
                    + triggerFacts.WindowHwnd
                    + NewCaseVisibilityObservation.FormatCorrelationFields(_excelInteropService, workbook));
                _workbookEventCoordinator.OnWindowActivate(triggerFacts);
            });
        }

        private void Application_SheetActivate(object sh)
        {
            _logger?.Debug("Application_SheetActivate", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetActivate), () => _sheetEventCoordinator?.OnSheetActivate(sh));
            _logger?.Debug("Application_SheetActivate", "returned.");
        }

        private void Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            _logger?.Debug("Application_SheetSelectionChange", "entry.");
            if (!(sh is Excel.Worksheet) || target == null)
            {
                _logger?.Debug("Application_SheetSelectionChange", "returned.");
                return;
            }

            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetSelectionChange), () =>
            {
                _sheetEventCoordinator?.OnSheetSelectionChange(sh, target);
            });
            _logger?.Debug("Application_SheetSelectionChange", "returned.");
        }

        private void Application_SheetChange(object sh, Excel.Range target)
        {
            _logger?.Debug("Application_SheetChange", "entry.");
            EventBoundaryGuard.Execute(_logger, nameof(Application_SheetChange), () => _sheetEventCoordinator?.OnSheetChange(sh, target));
            _logger?.Debug("Application_SheetChange", "returned.");
        }

        private void Application_AfterCalculate()
        {
            EventBoundaryGuard.Execute(_logger, nameof(Application_AfterCalculate), () => _sheetEventCoordinator?.OnAfterCalculate(_application));
            _logger?.Debug("Application_AfterCalculate", "EventBoundaryGuard.Execute returned.");
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook workbook, bool saveAsUi, ref bool cancel)
        {
            EventBoundaryGuard.ExecuteCancelable(_logger, nameof(Application_WorkbookBeforeSave), ref cancel, HandleBeforeSave);

            void HandleBeforeSave(ref bool innerCancel)
            {
                _logger?.Info(
                    "Excel WorkbookBeforeSave fired. workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook))
                    + ", saveAsUi="
                    + saveAsUi.ToString()
                    + ", cancel="
                    + innerCancel.ToString());
                _kernelWorkbookLifecycleService?.HandleWorkbookBeforeSave(workbook, saveAsUi, ref innerCancel);
            }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            EventBoundaryGuard.ExecuteCancelable(_logger, nameof(Application_WorkbookBeforeClose), ref cancel, HandleBeforeClose);

            void HandleBeforeClose(ref bool innerCancel)
            {
                if (_workbookLifecycleCoordinator != null)
                {
                    _workbookLifecycleCoordinator.OnWorkbookBeforeClose(workbook, ref innerCancel);
                }
            }
        }

        private WindowActivateTaskPaneTriggerFacts CaptureWindowActivateTaskPaneTriggerFacts(
            Excel.Workbook workbook,
            Excel.Window window,
            string captureOwner)
        {
            return new WindowActivateTaskPaneTriggerFacts(
                workbook,
                window,
                FormatWorkbookDescriptor(workbook),
                FormatWindowDescriptor(window),
                FormatActiveExcelState(),
                SafeWorkbookFullName(workbook),
                SafeWindowHwnd(window),
                captureOwner);
        }

        private void EnsureKernelFlickerTraceForWorkbookOpen(Excel.Workbook workbook)
        {
            if (!IsKernelWorkbookSafe(workbook))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(KernelFlickerTraceContext.CurrentTraceId))
            {
                return;
            }

            KernelFlickerTraceContext.BeginNewTrace();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=trace-begin trigger=WorkbookOpenKernelDetection workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState());
        }

        private bool IsKernelWorkbookSafe(Excel.Workbook workbook)
        {
            try
            {
                return workbook != null && _kernelWorkbookService != null && _kernelWorkbookService.IsKernelWorkbook(workbook);
            }
            catch
            {
                return false;
            }
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook);
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
