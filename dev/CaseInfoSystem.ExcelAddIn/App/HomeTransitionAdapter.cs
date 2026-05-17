using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class HomeTransitionAdapter
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Logger _logger;
        private readonly AddInRuntimeExecutionDiagnosticsService _runtimeExecutionDiagnosticsService;
        private readonly AddInExecutionBoundaryCoordinator _executionBoundaryCoordinator;
        private readonly TaskPaneEntryAdapter _taskPaneEntryAdapter;
        private readonly Func<string> _formatActiveExcelState;
        private KernelHomeFormHost _kernelHomeFormHost;
        private KernelWorkbookService _kernelWorkbookService;
        private KernelWorkbookAvailabilityService _kernelWorkbookAvailabilityService;
        private KernelHomeCasePaneSuppressionCoordinator _kernelHomeCasePaneSuppressionCoordinator;
        private ExternalWorkbookDetectionService _externalWorkbookDetectionService;
        private TaskPaneManager _taskPaneManager;
        private ExcelInteropService _excelInteropService;

        internal HomeTransitionAdapter(
            Logger logger,
            AddInRuntimeExecutionDiagnosticsService runtimeExecutionDiagnosticsService,
            AddInExecutionBoundaryCoordinator executionBoundaryCoordinator,
            TaskPaneEntryAdapter taskPaneEntryAdapter,
            Func<string> formatActiveExcelState)
        {
            _logger = logger;
            _runtimeExecutionDiagnosticsService = runtimeExecutionDiagnosticsService;
            _executionBoundaryCoordinator = executionBoundaryCoordinator;
            _taskPaneEntryAdapter = taskPaneEntryAdapter ?? throw new ArgumentNullException(nameof(taskPaneEntryAdapter));
            _formatActiveExcelState = formatActiveExcelState;
        }

        internal void Configure(
            KernelHomeFormHost kernelHomeFormHost,
            KernelWorkbookService kernelWorkbookService,
            KernelWorkbookAvailabilityService kernelWorkbookAvailabilityService,
            KernelHomeCasePaneSuppressionCoordinator kernelHomeCasePaneSuppressionCoordinator,
            ExternalWorkbookDetectionService externalWorkbookDetectionService,
            TaskPaneManager taskPaneManager,
            ExcelInteropService excelInteropService)
        {
            _kernelHomeFormHost = kernelHomeFormHost;
            _kernelWorkbookService = kernelWorkbookService;
            _kernelWorkbookAvailabilityService = kernelWorkbookAvailabilityService;
            _kernelHomeCasePaneSuppressionCoordinator = kernelHomeCasePaneSuppressionCoordinator;
            _externalWorkbookDetectionService = externalWorkbookDetectionService;
            _taskPaneManager = taskPaneManager;
            _excelInteropService = excelInteropService;
        }

        internal KernelHomeForm GetKernelHomeForm()
        {
            return _kernelHomeFormHost == null ? null : _kernelHomeFormHost.Current;
        }

        internal void ShowKernelHomePlaceholder(bool clearBindingOnNewSession = false)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.ResetKernelHomeExternalCloseRequested();

            _kernelHomeFormHost.GetOrCreate(clearBindingOnNewSession);
            _taskPaneManager?.HideKernelPanes();
            _kernelHomeFormHost.ReloadCurrent();

            _runtimeExecutionDiagnosticsService?.Trace("ShowKernelHomePlaceholder");
            _kernelWorkbookService.PrepareForHomeDisplayFromSheet();
            _kernelWorkbookService.EnsureHomeDisplayHidden("ThisAddIn.ShowKernelHomePlaceholder.BeforeShow");

            _kernelHomeFormHost.ShowAndActivateCurrent();
        }

        internal void ShowKernelHomeFromKernelCommand()
        {
            ShowKernelHomePlaceholderWithExternalWorkbookSuppression("KernelCommandService.OpenHome");
        }

        internal void ShowKernelHomePlaceholderWithExternalWorkbookSuppression(string reason)
        {
            ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(reason, clearBindingOnNewSession: false);
        }

        internal void ShowKernelHomePlaceholderWithExternalWorkbookSuppressionForNewSession(string reason)
        {
            ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(reason, clearBindingOnNewSession: true);
        }

        internal bool ShowKernelSheetAndRefreshPaneFromHome(WorkbookContext context, string sheetCodeName, string reason, out Excel.Workbook displayedWorkbook)
        {
            displayedWorkbook = null;
            if (context == null)
            {
                _logger?.Warn(
                    "ShowKernelSheetAndRefreshPaneFromHome skipped because workbook context was not available. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            Excel.Workbook resolvedDisplayedWorkbook = _kernelWorkbookService.ResolveKernelWorkbook(context);
            if (resolvedDisplayedWorkbook == null)
            {
                _logger?.Warn(
                    "ShowKernelSheetAndRefreshPaneFromHome skipped because bound kernel workbook could not be resolved. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            KernelFlickerTraceContext.BeginNewTrace();
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=trace-begin trigger=ShowKernelSheetAndRefreshPaneFromHomeBoundContext traceOriginReason="
                + (reason ?? string.Empty)
                + ", sheetCodeName="
                + (sheetCodeName ?? string.Empty)
                + ", workbook="
                + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(resolvedDisplayedWorkbook))
                + ", activeState="
                + FormatActiveExcelState());
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            bool shouldSuspendScreenUpdating = !string.IsNullOrWhiteSpace(reason)
                && reason.IndexOf("KernelHomeForm.OpenSheet", StringComparison.OrdinalIgnoreCase) >= 0;
            bool shown = false;
            Action performTransition = () =>
            {
                HideKernelHomePlaceholder();
                shown = _kernelWorkbookService.TryShowSheetByCodeName(context, sheetCodeName, reason);
                _logger?.Info("[Transition] bound-context sheet shown=" + shown + ", sheet=" + sheetCodeName);
                if (!shown)
                {
                    return;
                }

                _taskPaneEntryAdapter.RefreshTaskPane(reason, resolvedDisplayedWorkbook, null);
            };

            if (shouldSuspendScreenUpdating)
            {
                _executionBoundaryCoordinator.Execute(performTransition);
            }
            else
            {
                performTransition();
            }

            if (!shown)
            {
                _logger?.Info(
                    "ShowKernelSheetAndRefreshPaneFromHome aborted because target sheet could not be shown. reason="
                    + (reason ?? string.Empty)
                    + ", sheetCodeName="
                    + (sheetCodeName ?? string.Empty));
                return false;
            }

            displayedWorkbook = resolvedDisplayedWorkbook;
            return true;
        }

        internal void HandleKernelWorkbookBecameAvailable(string eventName, Excel.Workbook workbook)
        {
            _kernelWorkbookAvailabilityService?.Handle(eventName, workbook, GetKernelHomeForm());
        }

        internal void HandleExternalWorkbookDetected(Excel.Workbook workbook, string eventName)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.HandleExternalWorkbookDetected(
                _externalWorkbookDetectionService,
                workbook,
                eventName,
                GetKernelHomeForm());
        }

        internal void SuppressUpcomingKernelHomeDisplay(string reason, bool suppressOnOpen, bool suppressOnActivate)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen, suppressOnActivate);
        }

        internal bool ShouldSuppressKernelHomeDisplay(string eventName)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null && _kernelHomeCasePaneSuppressionCoordinator.ShouldSuppressKernelHomeDisplay(eventName);
        }

        internal void SuppressUpcomingCasePaneActivationRefresh(string workbookFullName, string reason)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.SuppressUpcomingCasePaneActivationRefresh(workbookFullName, reason);
        }

        internal void BeginCaseWorkbookActivateProtection(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            _kernelHomeCasePaneSuppressionCoordinator?.BeginCaseWorkbookActivateProtection(workbook, window, reason);
        }

        internal bool ShouldIgnoreWorkbookActivateDuringCaseProtection(Excel.Workbook workbook)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldIgnoreWorkbookActivateDuringProtection(workbook);
        }

        internal bool ShouldIgnoreWindowActivateDuringCaseProtection(Excel.Workbook workbook, Excel.Window window)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldIgnoreWindowActivateDuringProtection(workbook, window);
        }

        internal bool ShouldIgnoreTaskPaneRefreshDuringCaseProtection(string reason, Excel.Workbook workbook, Excel.Window window)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldIgnoreTaskPaneRefreshDuringProtection(reason, workbook, window);
        }

        internal bool ShouldSuppressCasePaneRefresh(string eventName, Excel.Workbook workbook)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.ShouldSuppressCasePaneRefresh(eventName, workbook);
        }

        internal bool IsKernelHomeSuppressionActive(string eventName, bool consume)
        {
            return _kernelHomeCasePaneSuppressionCoordinator != null
                && _kernelHomeCasePaneSuppressionCoordinator.IsKernelHomeSuppressionActive(eventName, consume);
        }

        internal void CloseOnShutdown()
        {
            _kernelHomeFormHost?.CloseOnShutdown();
        }

        private void ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(string reason, bool clearBindingOnNewSession)
        {
            SuppressUpcomingKernelHomeDisplay(reason, suppressOnOpen: false, suppressOnActivate: true);
            ShowKernelHomePlaceholder(clearBindingOnNewSession);
        }

        private void HideKernelHomePlaceholder()
        {
            _kernelHomeFormHost?.HideCurrent();
        }

        private string FormatActiveExcelState()
        {
            return _formatActiveExcelState == null ? string.Empty : _formatActiveExcelState();
        }
    }
}
