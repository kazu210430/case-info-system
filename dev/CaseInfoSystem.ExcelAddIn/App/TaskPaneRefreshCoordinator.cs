using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshCoordinator
    {
        private readonly WorkbookSessionService _workbookSessionService;
        private readonly TaskPaneManager _taskPaneManager;
        private readonly ExcelWindowRecoveryService _excelWindowRecoveryService;
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Action _scheduleWordWarmup;

        internal TaskPaneRefreshCoordinator(
            WorkbookSessionService workbookSessionService,
            TaskPaneManager taskPaneManager,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            Logger logger,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Action scheduleWordWarmup)
        {
            _workbookSessionService = workbookSessionService;
            _taskPaneManager = taskPaneManager;
            _excelWindowRecoveryService = excelWindowRecoveryService;
            _logger = logger;
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow;
            _scheduleWordWarmup = scheduleWordWarmup;
        }

        internal TaskPaneRefreshAttemptResult TryRefreshTaskPane(string reason, Excel.Workbook workbook, Excel.Window window, KernelHomeForm kernelHomeForm, int taskPaneRefreshSuppressionCount)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            if (!CanExecuteTaskPaneRefresh(reason, stopwatch, taskPaneRefreshSuppressionCount))
            {
                return TaskPaneRefreshAttemptResult.Skipped();
            }

            // recovery は単なる UI 修復ではなく、後続の context 解決の前提調整。
            // ActiveWindow / 可視 window / UI 状態を「解決可能な状態」に整える段階であり、
            // ここでは対象 window も context もまだ決定しない。
            if ((kernelHomeForm == null || kernelHomeForm.IsDisposed || !kernelHomeForm.Visible)
                && _excelWindowRecoveryService != null)
            {
                if (workbook != null)
                {
                    _excelWindowRecoveryService.TryRecoverWorkbookWindow(workbook, "TryRefreshTaskPane." + (reason ?? string.Empty), bringToFront: false);
                }
                else
                {
                    _excelWindowRecoveryService.TryRecoverActiveWorkbookWindow("TryRefreshTaskPane." + (reason ?? string.Empty), bringToFront: false);
                }
            }

            // recovery の次に、workbook 指定時の pane 対象 window をここで確定させる。
            // この段階で「どの window を対象にするか」が決まり、
            // 後続の context 生成はこの確定済み window に依存する。
            window = EnsurePaneWindowForWorkbook(workbook, window, reason, stopwatch);

            // ここでは pane 更新に使う実行 context を生成する。
            // context.Window は pane の表示先 / hide 先の正本だが、
            // この段階ではまだ表示採否や hide 判断は行わない。
            WorkbookContext context = workbook == null
                ? _workbookSessionService.ResolveActiveContext(reason)
                : _workbookSessionService.ResolveContext(workbook, window, reason);
            _logger?.Info("TryRefreshTaskPane context resolved. reason=" + (reason ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString() + ", role=" + (context == null ? string.Empty : context.Role.ToString()));

            // ここでは生成済み context を pane 対象として採用するかを調停する。
            // 対象外なら context を使わず、必要に応じて hide もここで判断する。
            // 生成と受理は分離されており、context 生成 → 受理判定の順で直列に進む。
            if (!TryAcceptTaskPaneContext(context, window, kernelHomeForm))
            {
                return TaskPaneRefreshAttemptResult.ContextRejected();
            }

            // 受理された context を使って pane UI へ反映し、
            // CASE 成功時の warmup もこの最終段でまとめて扱う。
            bool refreshed = TryRefreshPaneAndScheduleWarmup(context, reason, stopwatch);
            if (refreshed && window != null && _excelWindowRecoveryService != null)
            {
                if (workbook != null)
                {
                    _excelWindowRecoveryService.TryRecoverWorkbookWindow(workbook, "TryRefreshTaskPane.PostRefresh." + (reason ?? string.Empty), bringToFront: true);
                }
                else
                {
                    _excelWindowRecoveryService.TryRecoverActiveWorkbookWindow("TryRefreshTaskPane.PostRefresh." + (reason ?? string.Empty), bringToFront: true);
                }
            }

            return refreshed
                ? TaskPaneRefreshAttemptResult.Succeeded()
                : TaskPaneRefreshAttemptResult.Failed();
        }

        private bool CanExecuteTaskPaneRefresh(string reason, Stopwatch stopwatch, int taskPaneRefreshSuppressionCount)
        {
            if (_workbookSessionService == null || _taskPaneManager == null)
            {
                return false;
            }

            if (taskPaneRefreshSuppressionCount > 0)
            {
                _logger?.Info(
                    "TryRefreshTaskPane suppressed. reason="
                    + (reason ?? string.Empty)
                    + ", elapsedMs="
                    + stopwatch.ElapsedMilliseconds.ToString()
                    + ", suppressionCount="
                    + taskPaneRefreshSuppressionCount.ToString());
                return false;
            }

            return true;
        }

        private Excel.Window EnsurePaneWindowForWorkbook(Excel.Workbook workbook, Excel.Window window, string reason, Stopwatch stopwatch)
        {
            // workbook は既に特定できているが、pane 操作に使う window が未確定な場合は、
            // その workbook 専用に window を解決する。
            // ここで補完した window は、後続の context 解決と hide/show 対象 window の決定に使われる。
            if (workbook != null && window == null)
            {
                window = _resolveWorkbookPaneWindow(workbook, reason, false);
                _logger?.Info("TryRefreshTaskPane window resolved. reason=" + (reason ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString() + ", hasWindow=" + (window != null).ToString());
            }

            return window;
        }

        private bool TryAcceptTaskPaneContext(WorkbookContext context, Excel.Window window, KernelHomeForm kernelHomeForm)
        {
            if (kernelHomeForm != null && !kernelHomeForm.IsDisposed && kernelHomeForm.Visible)
            {
                if (context != null && context.Role == WorkbookRole.Kernel)
                {
                    _taskPaneManager.HideKernelPanes();
                    return false;
                }
            }

            if (!_workbookSessionService.ShouldHandleContext(context))
            {
                if (context != null)
                {
                    _taskPaneManager.HidePaneForWindow(context.Window);
                }
                else if (window != null)
                {
                    _taskPaneManager.HidePaneForWindow(window);
                }

                return false;
            }

            return true;
        }

        private bool TryRefreshPaneAndScheduleWarmup(WorkbookContext context, string reason, Stopwatch stopwatch)
        {
            bool refreshed = _taskPaneManager.RefreshPane(context, reason);
            _logger?.Info("TryRefreshTaskPane refresh completed. reason=" + (reason ?? string.Empty) + ", elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString() + ", refreshed=" + refreshed.ToString());
            if (refreshed && context != null && context.Role == WorkbookRole.Case)
            {
                _scheduleWordWarmup();
            }

            return refreshed;
        }
    }
}
