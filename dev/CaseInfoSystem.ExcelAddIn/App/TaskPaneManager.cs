using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed partial class TaskPaneManager
    {
        private readonly ExcelInteropService _excelInteropService;
        private readonly CasePaneSnapshotRenderService _casePaneSnapshotRenderService;
        private readonly Logger _logger;
        private readonly Dictionary<string, TaskPaneHost> _hostsByWindowKey;
        private TaskPaneHostLifecycleService _taskPaneHostLifecycleService;
        private TaskPaneDisplayCoordinator _taskPaneDisplayCoordinator;
        private TaskPaneHostFlowService _taskPaneHostFlowService;
        private CasePaneCacheRefreshNotificationService _casePaneCacheRefreshNotificationService;

        private TaskPaneManager(
            ExcelInteropService excelInteropService,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _casePaneSnapshotRenderService = casePaneSnapshotRenderService ?? throw new ArgumentNullException(nameof(casePaneSnapshotRenderService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
        }

        private TaskPaneManager(Logger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _excelInteropService = null;
            _casePaneSnapshotRenderService = null;
            _hostsByWindowKey = new Dictionary<string, TaskPaneHost>(StringComparer.OrdinalIgnoreCase);
        }

        private void AttachRuntimeGraph(TaskPaneManagerRuntimeGraph runtimeGraph)
        {
            if (runtimeGraph == null)
            {
                throw new ArgumentNullException(nameof(runtimeGraph));
            }

            if (_taskPaneHostFlowService != null)
            {
                throw new InvalidOperationException("TaskPaneManager runtime graph is already attached.");
            }

            // Attach only runtime-consumed collaborators here. Registry ownership stays below lifecycle/compose layers.
            _casePaneCacheRefreshNotificationService = runtimeGraph.CasePaneCacheRefreshNotificationService ?? throw new ArgumentException("Case pane cache notification service is required.", nameof(runtimeGraph));
            _taskPaneHostLifecycleService = runtimeGraph.TaskPaneHostLifecycleService ?? throw new ArgumentException("Task pane host lifecycle service is required.", nameof(runtimeGraph));
            _taskPaneDisplayCoordinator = runtimeGraph.TaskPaneDisplayCoordinator ?? throw new ArgumentException("Task pane display coordinator is required.", nameof(runtimeGraph));
            _taskPaneHostFlowService = runtimeGraph.TaskPaneHostFlowService ?? throw new ArgumentException("Task pane host flow service is required.", nameof(runtimeGraph));
        }

        internal bool RefreshPane(WorkbookContext context, string reason)
        {
            return _taskPaneHostFlowService.RefreshPane(context, reason);
        }

        internal bool TryShowExistingPane(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            return _taskPaneDisplayCoordinator.TryShowExistingPane(_excelInteropService, workbook, window, reason);
        }

        internal TaskPaneDisplayEntryState EvaluateDisplayEntryState(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.EvaluateDisplayEntryState(_excelInteropService, workbook, window);
        }

        internal bool HasManagedPaneForWindow(Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.HasManagedPaneForWindow(window);
        }

        internal bool HasVisibleCasePaneForWorkbookWindow(Excel.Workbook workbook, Excel.Window window)
        {
            return _taskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(_excelInteropService, workbook, window);
        }

        // Display 制御 facade: 表示前調停と hide/show は display coordinator に委譲する。
        internal void HideAll()
        {
            _taskPaneDisplayCoordinator.HideAll();
        }

        internal void HideKernelPanes()
        {
            _taskPaneDisplayCoordinator.HideKernelPanes();
        }

        internal void HideAllExcept(string activeWindowKey)
        {
            _taskPaneDisplayCoordinator.HideAllExcept(activeWindowKey);
        }

        /// <summary>
        /// メソッド: host 表示前に、他ウィンドウ pane の表示状態を必要最小限だけ調整する。
        /// 引数: host - これから表示する host。
        /// 戻り値: なし。
        /// 副作用: CASE pane は複数窓で維持し、それ以外は従来どおり単一表示に寄せる。
        /// </summary>
        internal void PrepareHostsBeforeShow(TaskPaneHost host)
        {
            _taskPaneDisplayCoordinator.PrepareHostsBeforeShow(host);
        }

        // Host lifecycle facade: workbook / host 単位の register/remove/dispose は lifecycle service に委譲する。
        /// <summary>
        /// メソッド: 指定 workbook に紐づく pane だけを非表示にして破棄する。
        /// 引数: workbook - 対象 workbook。
        /// 戻り値: なし。
        /// 副作用: 対象 workbook の Host を破棄し、他 workbook の pane には影響しない。
        /// </summary>
        internal void RemoveWorkbookPanes(Excel.Workbook workbook)
        {
            _taskPaneHostLifecycleService.RemoveWorkbookPanes(workbook);
        }

        internal void RegisterHost(TaskPaneHost host)
        {
            _taskPaneHostLifecycleService.RegisterHost(host);
        }

        /// <summary>
        /// メソッド: 指定 window に紐づく pane だけを非表示にする。
        /// 引数: window - 対象 window。
        /// 戻り値: なし。
        /// 副作用: 対象 window の pane だけを非表示にする。
        /// </summary>
        internal void HidePaneForWindow(Excel.Window window)
        {
            _taskPaneDisplayCoordinator.HidePaneForWindow(window);
        }

        internal void DisposeAll()
        {
            _taskPaneHostLifecycleService.DisposeAll();
        }

        // Render 制御責務: role ごとの描画と signature 更新対象を分離し、再描画条件は上位から受け取る。
        private void RenderHost(TaskPaneHost host, WorkbookContext context, string reason)
        {
            host.WorkbookFullName = context.WorkbookFullName;

            if (host.Control is KernelNavigationControl kernelControl)
            {
                RenderKernelHost(kernelControl, context);
                return;
            }

            if (host.Control is AccountingNavigationControl accountingControl)
            {
                RenderAccountingHost(accountingControl, context);
                return;
            }

            if (host.Control is DocumentButtonsControl caseControl)
            {
                RenderCaseHost(caseControl, context, reason);
            }
        }

        private void RenderKernelHost(KernelNavigationControl kernelControl, WorkbookContext context)
        {
            _logger.Info("RenderHost start. role=Kernel, workbook=" + (context.WorkbookFullName ?? string.Empty));
            kernelControl.Render(KernelNavigationDefinitions.CreateForSheet(context.ActiveSheetCodeName));
            _logger.Info("RenderHost completed. role=Kernel, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        private void RenderAccountingHost(AccountingNavigationControl accountingControl, WorkbookContext context)
        {
            _logger.Info("RenderHost start. role=Accounting, workbook=" + (context.WorkbookFullName ?? string.Empty));
            accountingControl.Render(AccountingNavigationDefinitions.CreateForSheet(context.ActiveSheetCodeName));
            _logger.Info("RenderHost completed. role=Accounting, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        private void RenderCaseHost(DocumentButtonsControl caseControl, WorkbookContext context, string reason)
        {
            _logger.Info("RenderHost start. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
            bool? originalWorkbookSavedState = _casePaneCacheRefreshNotificationService.TryGetWorkbookSavedState(context.Workbook);
            CasePaneSnapshotRenderService.CasePaneSnapshotRenderResult renderResult = _casePaneSnapshotRenderService.Render(caseControl, context.Workbook);
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult = renderResult.BuildResult;
            string snapshotText = buildResult.SnapshotText;
            _logger.Info("RenderHost snapshot acquired. role=Case, length=" + snapshotText.Length.ToString());
            TaskPaneSnapshot snapshot = renderResult.Snapshot;
            _logger.Info("RenderHost snapshot parsed. role=Case, hasError=" + snapshot.HasError.ToString() + ", tabs=" + snapshot.Tabs.Count.ToString() + ", docs=" + snapshot.DocButtons.Count.ToString());
            _casePaneCacheRefreshNotificationService.NotifyCasePaneUpdatedIfNeeded(context.Workbook, reason, buildResult, originalWorkbookSavedState);
            _logger.Info("RenderHost completed. role=Case, workbook=" + (context.WorkbookFullName ?? string.Empty));
        }

        internal void PrepareTargetWindowForForcedRefresh(Excel.Window targetWindow)
        {
            _taskPaneDisplayCoordinator.PrepareTargetWindowForForcedRefresh(targetWindow);
        }

        private string FormatContextDescriptor(WorkbookContext context)
        {
            return TaskPaneManagerDiagnosticHelper.FormatContextDescriptor(
                context,
                context == null ? string.Empty : FormatWorkbookDescriptor(context.Workbook, context.WorkbookFullName),
                context == null ? string.Empty : TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor(context.Window));
        }

        private string FormatHostDescriptor(TaskPaneHost host)
        {
            return TaskPaneManagerDiagnosticHelper.FormatHostDescriptor(
                host,
                TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor(host == null ? null : host.Window));
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook)
        {
            return FormatWorkbookDescriptor(workbook, null);
        }

        private string FormatWorkbookDescriptor(Excel.Workbook workbook, string fallbackFullName)
        {
            return "full=\""
                + SafeWorkbookFullName(workbook, fallbackFullName)
                + "\",name=\""
                + SafeWorkbookShortName(workbook)
                + "\"";
        }

        private string SafeWorkbookFullName(Excel.Workbook workbook, string fallbackFullName)
        {
            string workbookFullName = workbook == null || _excelInteropService == null
                ? string.Empty
                : (_excelInteropService.GetWorkbookFullName(workbook) ?? string.Empty);
            return string.IsNullOrWhiteSpace(workbookFullName) ? (fallbackFullName ?? string.Empty) : workbookFullName;
        }

        private string SafeWorkbookShortName(Excel.Workbook workbook)
        {
            return workbook == null || _excelInteropService == null
                ? string.Empty
                : (_excelInteropService.GetWorkbookName(workbook) ?? string.Empty);
        }

        internal static string SafeGetWindowKey(Excel.Window window)
        {
            return TaskPaneManagerDiagnosticHelper.SafeGetWindowKey(window);
        }

        internal static string FormatWindowDescriptor(Excel.Window window)
        {
            return TaskPaneManagerDiagnosticHelper.FormatWindowDescriptor(window);
        }

        private static class TaskPaneManagerDiagnosticHelper
        {
            internal static bool IsCaseHost(TaskPaneHost host)
            {
                return host != null && host.Control is DocumentButtonsControl;
            }

            internal static bool IsKernelHost(TaskPaneHost host)
            {
                return host != null && host.Control is KernelNavigationControl;
            }

            internal static string SafeGetWindowKey(Excel.Window window)
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

            internal static string FormatContextDescriptor(WorkbookContext context, string workbookDescriptor, string windowDescriptor)
            {
                if (context == null)
                {
                    return "null";
                }

                return "role=\""
                    + context.Role.ToString()
                    + "\",workbook="
                    + (workbookDescriptor ?? string.Empty)
                    + ",window="
                    + (windowDescriptor ?? string.Empty)
                    + ",activeSheet=\""
                    + (context.ActiveSheetCodeName ?? string.Empty)
                    + "\"";
            }

            internal static string FormatHostDescriptor(TaskPaneHost host, string windowDescriptor)
            {
                if (host == null)
                {
                    return "null";
                }

                return "paneRole=\""
                    + GetPaneRoleName(host)
                    + "\",windowKey=\""
                    + (host.WindowKey ?? string.Empty)
                    + "\",workbookFullName=\""
                    + (host.WorkbookFullName ?? string.Empty)
                    + "\",window="
                    + (windowDescriptor ?? string.Empty);
            }

            internal static string FormatWindowDescriptor(Excel.Window window)
            {
                return "hwnd=\""
                    + SafeGetWindowKey(window)
                    + "\",caption=\""
                    + SafeWindowCaption(window)
                    + "\"";
            }

            private static string GetPaneRoleName(TaskPaneHost host)
            {
                if (host == null || host.Control == null)
                {
                    return "Unknown";
                }

                if (IsCaseHost(host))
                {
                    return "Case";
                }

                if (IsKernelHost(host))
                {
                    return "Kernel";
                }

                if (host.Control is AccountingNavigationControl)
                {
                    return "Accounting";
                }

                return host.Control.GetType().Name;
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
        }

        internal sealed class TaskPaneManagerTestHooks
        {
            internal Action<string, string> OnHideHost { get; set; }

            internal Func<string, string, bool> TryShowHost { get; set; }

            internal Action<string> OnCasePaneUpdatedNotification { get; set; }
        }
    }

    internal static class TaskPaneShowExistingPolicy
    {
        internal static bool ShouldShowExisting(bool hasExistingHost, bool isSameWorkbook, bool isRenderSignatureCurrent)
        {
            return hasExistingHost && isSameWorkbook && isRenderSignatureCurrent;
        }
    }

    internal static class TaskPaneShowWithRenderPolicy
    {
        internal static bool ShouldShowWithRender(bool hasExistingHost, bool isSameWorkbook, bool isRenderSignatureCurrent)
        {
            return !hasExistingHost || !isSameWorkbook || !isRenderSignatureCurrent;
        }
    }
}
