using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookDisplayService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly ExcelWindowRecoveryService _excelWindowRecoveryService;
        private readonly KernelCaseInteractionState _kernelCaseInteractionState;
        private readonly Logger _logger;
        private readonly KernelWorkbookBindingService _bindingService;
        private readonly KernelWorkbookService.KernelWorkbookServiceTestHooks _testHooks;
        private bool _isHomeDisplayPrepared;

        internal KernelWorkbookDisplayService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger,
            KernelWorkbookBindingService bindingService,
            KernelWorkbookService.KernelWorkbookServiceTestHooks testHooks = null)
        {
            _application = application;
            _excelInteropService = excelInteropService;
            _excelWindowRecoveryService = excelWindowRecoveryService;
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _bindingService = bindingService ?? throw new ArgumentNullException(nameof(bindingService));
            _testHooks = testHooks;
        }

        internal bool TryShowSheetByCodeName(WorkbookContext context, string sheetCodeName, string reason)
        {
            Excel.Workbook kernelWorkbook = _bindingService.ResolveKernelWorkbook(context);
            if (kernelWorkbook == null)
            {
                _logger.Warn("TryShowSheetByCodeName skipped because kernel workbook was not available. reason=" + (reason ?? string.Empty));
                return false;
            }

            PrepareWorkbookForSheetNavigation(kernelWorkbook, sheetCodeName);
            bool activated = _excelInteropService.ActivateWorkbook(kernelWorkbook);
            bool sheetActivated = activated && _excelInteropService.ActivateWorksheetByCodeName(kernelWorkbook, sheetCodeName);
            _logger.Info(
                "TryShowSheetByCodeName result=" + sheetActivated.ToString()
                + ", reason=" + (reason ?? string.Empty)
                + ", sheetCodeName=" + (sheetCodeName ?? string.Empty));
            return sheetActivated;
        }

        internal void PrepareForHomeDisplay()
        {
            if (_isHomeDisplayPrepared)
            {
                return;
            }

            ApplyHomeDisplayVisibilityCore("PrepareForHomeDisplay");
            _isHomeDisplayPrepared = true;
        }

        internal void PrepareForHomeDisplayFromSheet()
        {
            ApplyHomeDisplayVisibilityCore("PrepareForHomeDisplayFromSheet");
            _isHomeDisplayPrepared = true;
        }

        internal void CompleteHomeNavigation(bool showExcel)
        {
            ReleaseHomeDisplay(showExcel);
        }

        internal void EnsureHomeDisplayHidden(string triggerReason)
        {
            string caller = ResolveExternalCaller(typeof(KernelWorkbookDisplayService), typeof(KernelWorkbookService));
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=ensure-home-display-hidden-enter trigger="
                + (triggerReason ?? string.Empty)
                + ", caller="
                + caller
                + ", activeState="
                + FormatActiveExcelState()
                + ", isHomeDisplayPrepared="
                + _isHomeDisplayPrepared.ToString()
                + ", tracePresent="
                + (!string.IsNullOrWhiteSpace(KernelFlickerTraceContext.CurrentTraceId)).ToString());
            if (!_isHomeDisplayPrepared)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=ensure-home-display-hidden-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=skipped-not-prepared");
                return;
            }

            ApplyHomeDisplayVisibilityCore("EnsureHomeDisplayHidden|" + (triggerReason ?? string.Empty));
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=ensure-home-display-hidden-end trigger="
                + (triggerReason ?? string.Empty)
                + ", result=applied, activeState="
                + FormatActiveExcelState());
        }

        internal void ReleaseHomeDisplayCore(bool showExcel)
        {
            if (_testHooks != null && _testHooks.ReleaseHomeDisplay != null)
            {
                _testHooks.ReleaseHomeDisplay(showExcel);
                return;
            }

            ReleaseHomeDisplay(showExcel);
        }

        internal void DismissPreparedHomeDisplayStateCore(string reason)
        {
            if (_testHooks != null && _testHooks.DismissPreparedHomeDisplayState != null)
            {
                _testHooks.DismissPreparedHomeDisplayState(reason);
                return;
            }

            DismissPreparedHomeDisplayState(reason);
        }

        internal void ApplyHomeDisplayVisibilityCore(string triggerReason)
        {
            if (_testHooks != null && _testHooks.ApplyHomeDisplayVisibility != null)
            {
                _testHooks.ApplyHomeDisplayVisibility();
                return;
            }

            ApplyHomeDisplayVisibility(triggerReason);
        }

        internal void ConcealKernelWorkbookWindowsForCaseCreationCloseCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.ConcealKernelWorkbookWindowsForCaseCreationClose != null)
            {
                _testHooks.ConcealKernelWorkbookWindowsForCaseCreationClose(workbook);
                return;
            }

            ConcealKernelWorkbookWindowsForCaseCreationClose(workbook);
        }

        internal void ShowKernelWorkbookWindows(bool activateWorkbookWindow)
        {
            Excel.Workbook workbook = _bindingService.ResolveWorkbookForHomeDisplayOrClose("ShowKernelWorkbookWindows");
            if (workbook == null)
            {
                return;
            }

            try
            {
                _excelWindowRecoveryService.EnsureApplicationVisible("KernelWorkbookService.ShowKernelWorkbookWindows", _bindingService.GetWorkbookFullName(workbook));
                foreach (Excel.Window window in workbook.Windows)
                {
                    if (window != null)
                    {
                        window.Visible = true;
                        window.WindowState = Excel.XlWindowState.xlNormal;
                    }
                }

                if (activateWorkbookWindow)
                {
                    _excelWindowRecoveryService.TryRecoverWorkbookWindow(
                        workbook,
                        "KernelWorkbookService.ShowKernelWorkbookWindows",
                        bringToFront: true);
                }
            }
            catch (Exception ex)
            {
                _logger.Error("ShowKernelWorkbookWindows failed.", ex);
            }
        }

        internal void HideExcelMainWindow()
        {
            try
            {
                LogHideExcelMainWindowState("before");
                _excelWindowRecoveryService.HideApplicationWindow(
                    "KernelWorkbookService.HideExcelMainWindow",
                    SafeActiveWorkbookDescriptor());
                LogHideExcelMainWindowState("after");
            }
            catch
            {
            }
        }

        internal void ShowExcelMainWindow()
        {
            try
            {
                EnsureExcelApplicationVisible();
                _excelWindowRecoveryService.TryBringApplicationToForeground(
                    "KernelWorkbookService.ShowExcelMainWindow",
                    SafeActiveWorkbookDescriptor());
            }
            catch
            {
            }
        }

        internal void LogKernelFlickerTrace(string detail)
        {
            _logger.Info(KernelFlickerTracePrefix + " " + (detail ?? string.Empty));
        }

        internal string FormatActiveExcelState()
        {
            Excel.Workbook activeWorkbook = _excelInteropService == null ? null : _excelInteropService.GetActiveWorkbook();
            Excel.Window activeWindow = _excelInteropService == null ? null : _excelInteropService.GetActiveWindow();
            return "activeWorkbook=" + FormatWorkbookDescriptor(activeWorkbook) + ",activeWindow=" + FormatWindowDescriptor(activeWindow);
        }

        internal string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        internal string DescribeVisibleOtherWorkbookWindows(Excel.Workbook workbookToIgnore)
        {
            if (_application == null)
            {
                return "app-null";
            }

            List<string> descriptors = new List<string>();
            try
            {
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (workbook == null || ReferenceEquals(workbook, workbookToIgnore))
                    {
                        continue;
                    }

                    List<string> visibleWindows = new List<string>();
                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (SafeWindowVisibleValue(window))
                        {
                            visibleWindows.Add(FormatWindowDescriptor(window));
                        }
                    }

                    if (visibleWindows.Count > 0)
                    {
                        descriptors.Add(FormatWorkbookDescriptor(workbook) + " windows=[" + string.Join(" | ", visibleWindows) + "]");
                    }
                }
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }

            return descriptors.Count == 0 ? "none" : string.Join(" || ", descriptors);
        }

        private void PrepareWorkbookForSheetNavigation(Excel.Workbook workbook, string codeName)
        {
            if (_isHomeDisplayPrepared)
            {
                DismissPreparedHomeDisplayState("ShowSheetByCodeName:" + (codeName ?? string.Empty));
            }

            EnsureWorkbookVisible(workbook);
        }

        private void EnsureWorkbookVisible(Excel.Workbook workbook)
        {
            ReleaseHomeDisplay(true);
            bool shouldAvoidGlobalRestore = ShouldAvoidGlobalExcelWindowRestore();
            if (!shouldAvoidGlobalRestore)
            {
                _excelWindowRecoveryService.EnsureApplicationVisible("KernelWorkbookService.EnsureWorkbookVisible", _bindingService.GetWorkbookFullName(workbook));
            }

            _excelWindowRecoveryService.TryRecoverWorkbookWindow(
                workbook,
                "KernelWorkbookService.EnsureWorkbookVisible",
                bringToFront: !shouldAvoidGlobalRestore);
        }

        private void ReleaseHomeDisplay(bool showExcel)
        {
            if (!_isHomeDisplayPrepared)
            {
                return;
            }

            if (showExcel)
            {
                DispatchHomeDisplayReleaseBranchForShowingExcel();
            }

            _isHomeDisplayPrepared = false;
        }

        private void DispatchHomeDisplayReleaseBranchForShowingExcel()
        {
            KernelWorkbookHomeReleaseAction homeReleaseAction = ResolveHomeDisplayReleaseActionForShowingExcel();
            switch (homeReleaseAction)
            {
                case KernelWorkbookHomeReleaseAction.SkipRestore:
                    _logger.Info("ReleaseHomeDisplay skipped global Excel window restore to preserve other workbook layouts.");
                    return;

                case KernelWorkbookHomeReleaseAction.PromoteAndRestore:
                    ShowExcelMainWindow();
                    ShowKernelWorkbookWindows(activateWorkbookWindow: true);
                    return;

                case KernelWorkbookHomeReleaseAction.RestoreWithoutPromotion:
                default:
                    ShowKernelWorkbookWindows(activateWorkbookWindow: false);
                    return;
            }
        }

        private KernelWorkbookHomeReleaseAction ResolveHomeDisplayReleaseActionForShowingExcel()
        {
            bool shouldAvoidGlobalExcelWindowRestore = ShouldAvoidGlobalExcelWindowRestore();
            bool shouldPromoteKernelWindow = !shouldAvoidGlobalExcelWindowRestore
                && ShouldPromoteKernelWorkbookOnHomeRelease();
            return KernelWorkbookHomeReleaseFallbackPolicy.DecideHomeReleaseAction(
                shouldAvoidGlobalExcelWindowRestore: shouldAvoidGlobalExcelWindowRestore,
                shouldPromoteKernelWorkbook: shouldPromoteKernelWindow);
        }

        private void DismissPreparedHomeDisplayState(string reason)
        {
            if (!_isHomeDisplayPrepared)
            {
                return;
            }

            _isHomeDisplayPrepared = false;
            _logger.Info("DismissPreparedHomeDisplayState executed. reason=" + (reason ?? string.Empty));
        }

        private void ApplyHomeDisplayVisibility(string triggerReason)
        {
            Excel.Workbook kernelWorkbook = _bindingService.ResolveWorkbookForHomeDisplayOrClose("ApplyHomeDisplayVisibility");
            bool hasVisibleNonKernelWorkbook = _bindingService.HasVisibleNonKernelWorkbook();
            bool preserveOtherWorkbookWindowLayout = ShouldPreserveOtherWorkbookWindowLayout();
            string visibleNonKernelWindows = DescribeVisibleNonKernelWorkbookWindows();
            string kernelWindowTargets = DescribeWorkbookWindows(kernelWorkbook);
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=apply-home-display-enter trigger="
                + (triggerReason ?? string.Empty)
                + ", activeState="
                + FormatActiveExcelState()
                + ", hasVisibleNonKernelWorkbook="
                + hasVisibleNonKernelWorkbook.ToString()
                + ", preserveOtherWorkbookWindowLayout="
                + preserveOtherWorkbookWindowLayout.ToString()
                + ", visibleNonKernelWindows="
                + visibleNonKernelWindows
                + ", kernelWindowTargets="
                + kernelWindowTargets);

            Excel.Workbook activeWorkbook = null;
            bool isActiveWorkbookKernel = false;
            int visibleWorkbookCount = -1;
            try
            {
                activeWorkbook = _application == null ? null : _application.ActiveWorkbook;
                isActiveWorkbookKernel = activeWorkbook != null && _bindingService.IsKernelWorkbook(activeWorkbook);
                visibleWorkbookCount = CountVisibleWorkbooksSafe();
            }
            catch
            {
                activeWorkbook = null;
                isActiveWorkbookKernel = false;
                visibleWorkbookCount = -1;
            }

            KernelWorkbookHomeDisplayVisibilityAction visibilityAction = KernelWorkbookHomeDisplayVisibilityPolicy.DecideAction(
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook,
                isActiveWorkbookKernel: isActiveWorkbookKernel,
                visibleWorkbookCount: visibleWorkbookCount);

            if (visibilityAction == KernelWorkbookHomeDisplayVisibilityAction.MinimizeKernelWindows)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=apply-home-display-decision trigger="
                    + (triggerReason ?? string.Empty)
                    + ", decision=minimize-kernel-windows, reason=visible-non-kernel-workbook-detected, preserveOtherWorkbookWindowLayout="
                    + preserveOtherWorkbookWindowLayout.ToString()
                    + ", visibleNonKernelWindows="
                    + visibleNonKernelWindows
                    + ", kernelWindowTargets="
                    + kernelWindowTargets);
                if (!preserveOtherWorkbookWindowLayout)
                {
                    EnsureExcelApplicationVisible();
                }

                HideKernelWorkbookWindows("ApplyHomeDisplayVisibility:" + (triggerReason ?? string.Empty), kernelWorkbook);
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=apply-home-display-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=minimized-kernel-windows, visibleNonKernelWindows="
                    + visibleNonKernelWindows
                    + ", kernelWindowTargets="
                    + kernelWindowTargets);
                _logger.Info(
                    "ApplyHomeDisplayVisibility minimized kernel windows because a non-kernel workbook is visible. preserveOtherWindowLayout="
                    + preserveOtherWorkbookWindowLayout.ToString());
                return;
            }

            if (visibilityAction == KernelWorkbookHomeDisplayVisibilityAction.ConcealKernelWindowsAndHideExcelMainWindow)
            {
                Excel.Workbook workbookToConceal = kernelWorkbook ?? activeWorkbook;
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=apply-home-display-decision trigger="
                    + (triggerReason ?? string.Empty)
                    + ", decision=conceal-kernel-windows-and-hide-excel-main-window, reason=active-kernel-workbook-still-visible, visibleWorkbookCount="
                    + visibleWorkbookCount.ToString()
                    + ", activeWorkbook="
                    + FormatWorkbookDescriptor(activeWorkbook)
                    + ", concealTarget="
                    + FormatWorkbookDescriptor(workbookToConceal)
                    + ", kernelWindowTargets="
                    + kernelWindowTargets);
                ConcealKernelWorkbookWindowsForHomeDisplay(workbookToConceal, "ApplyHomeDisplayVisibility:" + (triggerReason ?? string.Empty));
                _logger.Info(
                    "ApplyHomeDisplayVisibility concealed kernel windows before hiding Excel main window because the active kernel workbook remained visible.");
            }

            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=apply-home-display-decision trigger="
                + (triggerReason ?? string.Empty)
                + ", decision="
                + (visibilityAction == KernelWorkbookHomeDisplayVisibilityAction.HideExcelMainWindowOnly
                    ? "hide-excel-main-window-only"
                    : "hide-excel-main-window")
                + ", reason=no-visible-non-kernel-workbook, kernelWindowTargets="
                + kernelWindowTargets);
            HideExcelMainWindow();
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=apply-home-display-end trigger="
                + (triggerReason ?? string.Empty)
                + ", result=excel-main-window-hidden, activeState="
                + FormatActiveExcelState());
        }

        private bool ShouldPreserveOtherWorkbookWindowLayout()
        {
            return !_kernelCaseInteractionState.IsKernelCaseCreationFlowActive;
        }

        private bool ShouldAvoidGlobalExcelWindowRestore()
        {
            return KernelWorkbookWindowRestorePolicy.ShouldAvoidGlobalExcelWindowRestore(
                isKernelCaseCreationFlowActive: _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                hasVisibleNonKernelWorkbook: _bindingService.HasVisibleNonKernelWorkbook());
        }

        private void LogHideExcelMainWindowState(string stage)
        {
            _logger.Info(
                "HideExcelMainWindow state. stage="
                + (stage ?? string.Empty)
                + ", applicationVisible="
                + SafeApplicationVisible()
                + ", applicationHwnd="
                + SafeApplicationHwnd()
                + ", activeWorkbook="
                + SafeActiveWorkbookDescriptor()
                + ", activeWindow="
                + SafeActiveWindowDescriptor()
                + ", visibleWorkbookCount="
                + CountVisibleWorkbooksSafe().ToString());
        }

        private string SafeApplicationVisible()
        {
            try
            {
                return _application == null ? string.Empty : _application.Visible.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeApplicationHwnd()
        {
            try
            {
                return _application == null ? string.Empty : Convert.ToString(_application.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeActiveWorkbookDescriptor()
        {
            try
            {
                Excel.Workbook workbook = _application == null ? null : _application.ActiveWorkbook;
                return workbook == null ? string.Empty : _bindingService.GetWorkbookFullName(workbook);
            }
            catch
            {
                return string.Empty;
            }
        }

        private string SafeActiveWindowDescriptor()
        {
            try
            {
                Excel.Window window = _application == null ? null : _application.ActiveWindow;
                return window == null ? string.Empty : FormatWindowDescriptor(window);
            }
            catch
            {
                return string.Empty;
            }
        }

        private int CountVisibleWorkbooksSafe()
        {
            try
            {
                int count = 0;
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (workbook == null || workbook.Windows == null)
                    {
                        continue;
                    }

                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (window != null && window.Visible)
                        {
                            count++;
                            break;
                        }
                    }
                }

                return count;
            }
            catch
            {
                return -1;
            }
        }

        private void HideKernelWorkbookWindows(Excel.Workbook workbook)
        {
            HideKernelWorkbookWindows("HideKernelWorkbookWindows.DirectWorkbook", workbook);
        }

        private void HideKernelWorkbookWindows(string triggerReason, Excel.Workbook workbook)
        {
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=hide-kernel-windows-enter trigger="
                + (triggerReason ?? string.Empty)
                + ", targetWorkbook="
                + FormatWorkbookDescriptor(workbook)
                + ", activeState="
                + FormatActiveExcelState()
                + ", targets="
                + DescribeWorkbookWindows(workbook));
            if (workbook == null)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=hide-kernel-windows-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=skipped-null-workbook");
                return;
            }

            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                int minimizedCount = 0;
                int failedCount = 0;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        string beforeState = FormatWindowDescriptor(window);
                        LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=minimize-window-start trigger="
                            + (triggerReason ?? string.Empty)
                            + ", index="
                            + index.ToString()
                            + ", workbook="
                            + FormatWorkbookDescriptor(workbook)
                            + ", window="
                            + beforeState);
                        if (window != null)
                        {
                            bool isVisible = false;
                            try
                            {
                                isVisible = window.Visible;
                            }
                            catch
                            {
                            }

                            if (!isVisible)
                            {
                                LogKernelFlickerTrace(
                                    "source=KernelWorkbookService action=minimize-window-end trigger="
                                    + (triggerReason ?? string.Empty)
                                    + ", index="
                                    + index.ToString()
                                    + ", result=skipped-already-invisible, workbook="
                                    + FormatWorkbookDescriptor(workbook)
                                    + ", window="
                                    + beforeState);
                                continue;
                            }

                            window.WindowState = Excel.XlWindowState.xlMinimized;
                            minimizedCount++;
                            LogKernelFlickerTrace(
                                "source=KernelWorkbookService action=minimize-window-end trigger="
                                + (triggerReason ?? string.Empty)
                                + ", index="
                                + index.ToString()
                                + ", result=success, workbook="
                                + FormatWorkbookDescriptor(workbook)
                                + ", windowBefore="
                                + beforeState
                                + ", windowAfter="
                                + FormatWindowDescriptor(window));
                        }
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=minimize-window-end trigger="
                            + (triggerReason ?? string.Empty)
                            + ", index="
                            + index.ToString()
                            + ", result=failed, workbook="
                            + FormatWorkbookDescriptor(workbook)
                            + ", window="
                            + FormatWindowDescriptor(window)
                            + ", exceptionType="
                            + ex.GetType().Name
                            + ", exceptionMessage="
                            + (ex.Message ?? string.Empty));
                        _logger.Error("HideKernelWorkbookWindows window minimize failed. index=" + index.ToString(), ex);
                    }
                }

                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=hide-kernel-windows-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=completed, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", totalTargets="
                    + windowCount.ToString()
                    + ", minimizedCount="
                    + minimizedCount.ToString()
                    + ", failedCount="
                    + failedCount.ToString());
            }
            catch (Exception ex)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=hide-kernel-windows-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", result=failed, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", exceptionType="
                    + ex.GetType().Name
                    + ", exceptionMessage="
                    + (ex.Message ?? string.Empty));
                _logger.Error("HideKernelWorkbookWindows failed.", ex);
            }
        }

        private bool ShouldPromoteKernelWorkbookOnHomeRelease()
        {
            Excel.Workbook activeWorkbook = null;
            try
            {
                activeWorkbook = _application.ActiveWorkbook;
            }
            catch (Exception ex)
            {
                _logger.Error("ShouldPromoteKernelWorkbookOnHomeRelease failed to resolve ActiveWorkbook.", ex);
            }

            bool hasActiveWorkbook = activeWorkbook != null;
            bool isActiveWorkbookKernel = hasActiveWorkbook && _bindingService.IsKernelWorkbook(activeWorkbook);
            bool hasVisibleNonKernelWorkbook = !hasActiveWorkbook && _bindingService.HasVisibleNonKernelWorkbook();
            bool shouldPromoteKernelWorkbook = KernelWorkbookPromotionPolicy.ShouldPromoteKernelWorkbookOnHomeRelease(
                isKernelCaseCreationFlowActive: _kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                hasActiveWorkbook: hasActiveWorkbook,
                isActiveWorkbookKernel: isActiveWorkbookKernel,
                hasVisibleNonKernelWorkbook: hasVisibleNonKernelWorkbook);

            if (shouldPromoteKernelWorkbook)
            {
                return true;
            }

            _logger.Info(
                "Kernel workbook promotion skipped to preserve active non-kernel workbook. activeWorkbook="
                + _bindingService.GetWorkbookFullName(activeWorkbook));
            return false;
        }

        private void EnsureExcelApplicationVisible()
        {
            try
            {
                _excelWindowRecoveryService.ShowApplicationWindow(
                    "KernelWorkbookService.EnsureExcelApplicationVisible",
                    SafeActiveWorkbookDescriptor());
            }
            catch
            {
            }
        }

        private void ConcealKernelWorkbookWindowsForHomeDisplay(Excel.Workbook workbook, string triggerReason)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        if (window == null)
                        {
                            continue;
                        }

                        SetKernelWindowVisibleFalse(
                            workbook,
                            window,
                            index,
                            "ConcealKernelWorkbookWindowsForHomeDisplay|" + (triggerReason ?? string.Empty));
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("ConcealKernelWorkbookWindowsForHomeDisplay window conceal failed. index=" + index.ToString(), ex);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("ConcealKernelWorkbookWindowsForHomeDisplay failed.", ex);
            }
        }

        private void ConcealKernelWorkbookWindowsForCaseCreationClose(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = workbook.Windows[index];
                        if (window == null)
                        {
                            continue;
                        }

                        SetKernelWindowVisibleFalse(
                            workbook,
                            window,
                            index,
                            "ConcealKernelWorkbookWindowsForCaseCreationClose");
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("ConcealKernelWorkbookWindowsForCaseCreationClose window conceal failed. index=" + index.ToString(), ex);
                    }
                }

                _logger.Info("ConcealKernelWorkbookWindowsForCaseCreationClose completed. workbook=" + _bindingService.GetWorkbookFullName(workbook));
            }
            catch (Exception ex)
            {
                _logger.Error("ConcealKernelWorkbookWindowsForCaseCreationClose failed.", ex);
            }
        }

        private void SetKernelWindowVisibleFalse(Excel.Workbook workbook, Excel.Window window, int index, string triggerReason)
        {
            string caller = ResolveExternalCaller(typeof(KernelWorkbookDisplayService), typeof(KernelWorkbookService));
            string beforeState = FormatWindowDescriptor(window);
            LogKernelFlickerTrace(
                "source=KernelWorkbookService action=set-window-visible-false-start trigger="
                + (triggerReason ?? string.Empty)
                + ", caller="
                + caller
                + ", index="
                + index.ToString()
                + ", workbook="
                + FormatWorkbookDescriptor(workbook)
                + ", windowBefore="
                + beforeState
                + ", activeState="
                + FormatActiveExcelState());

            try
            {
                if (window == null)
                {
                    LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=set-window-visible-false-end trigger="
                        + (triggerReason ?? string.Empty)
                        + ", caller="
                        + caller
                        + ", index="
                        + index.ToString()
                        + ", result=skipped-null-window, workbook="
                        + FormatWorkbookDescriptor(workbook));
                    return;
                }

                window.Visible = false;
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=set-window-visible-false-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", caller="
                    + caller
                    + ", index="
                    + index.ToString()
                    + ", result=success, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", windowBefore="
                    + beforeState
                    + ", windowAfter="
                    + FormatWindowDescriptor(window));
            }
            catch (Exception ex)
            {
                LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=set-window-visible-false-end trigger="
                    + (triggerReason ?? string.Empty)
                    + ", caller="
                    + caller
                    + ", index="
                    + index.ToString()
                    + ", result=failed, workbook="
                    + FormatWorkbookDescriptor(workbook)
                    + ", window="
                    + FormatWindowDescriptor(window)
                    + ", exceptionType="
                    + ex.GetType().Name
                    + ", exceptionMessage="
                    + (ex.Message ?? string.Empty));
                throw;
            }
        }

        private string DescribeVisibleNonKernelWorkbookWindows()
        {
            if (_application == null)
            {
                return "app-null";
            }

            List<string> descriptors = new List<string>();
            try
            {
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (workbook == null || _bindingService.IsKernelWorkbook(workbook))
                    {
                        continue;
                    }

                    List<string> visibleWindows = new List<string>();
                    foreach (Excel.Window window in workbook.Windows)
                    {
                        if (SafeWindowVisibleValue(window))
                        {
                            visibleWindows.Add(FormatWindowDescriptor(window));
                        }
                    }

                    if (visibleWindows.Count > 0)
                    {
                        descriptors.Add(FormatWorkbookDescriptor(workbook) + " windows=[" + string.Join(" | ", visibleWindows) + "]");
                    }
                }
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }

            return descriptors.Count == 0 ? "none" : string.Join(" || ", descriptors);
        }

        private string DescribeWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return "none";
            }

            List<string> descriptors = new List<string>();
            try
            {
                int windowCount = workbook.Windows == null ? 0 : workbook.Windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = workbook.Windows[index];
                    descriptors.Add("index=" + index.ToString() + "," + FormatWindowDescriptor(window));
                }
            }
            catch (Exception ex)
            {
                return "enumeration-failed:" + ex.GetType().Name;
            }

            return descriptors.Count == 0 ? "none" : string.Join(" | ", descriptors);
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook)
        {
            return _bindingService.GetWorkbookFullName(workbook);
        }

        private string SafeWorkbookName(Excel.Workbook workbook)
        {
            return _bindingService.GetWorkbookName(workbook);
        }

        private string FormatWindowDescriptor(Excel.Window window)
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

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd) ?? string.Empty;
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

        private static string ResolveExternalCaller(params Type[] typesToSkip)
        {
            try
            {
                StackTrace stackTrace = new StackTrace(skipFrames: 1, fNeedFileInfo: false);
                StackFrame[] frames = stackTrace.GetFrames();
                if (frames == null)
                {
                    return string.Empty;
                }

                foreach (StackFrame frame in frames)
                {
                    var method = frame.GetMethod();
                    if (method == null)
                    {
                        continue;
                    }

                    Type declaringType = method.DeclaringType;
                    bool shouldSkip = false;
                    if (typesToSkip != null)
                    {
                        foreach (Type typeToSkip in typesToSkip)
                        {
                            if (declaringType == typeToSkip)
                            {
                                shouldSkip = true;
                                break;
                            }
                        }
                    }

                    if (shouldSkip)
                    {
                        continue;
                    }

                    string typeName = declaringType == null ? string.Empty : declaringType.FullName ?? string.Empty;
                    return string.IsNullOrWhiteSpace(typeName) ? method.Name : typeName + "." + method.Name;
                }
            }
            catch
            {
            }

            return string.Empty;
        }
    }
}
