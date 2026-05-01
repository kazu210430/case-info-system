using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class CasePaneCacheRefreshNotificationService
    {
        private const string ProductTitle = "案件情報System";

        private readonly Logger _logger;
        private readonly Func<Excel.Workbook, string> _safeGetWorkbookName;
        private readonly Action<string> _onCasePaneUpdatedNotification;

        internal CasePaneCacheRefreshNotificationService(
            Logger logger,
            Func<Excel.Workbook, string> safeGetWorkbookName,
            Action<string> onCasePaneUpdatedNotification = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _safeGetWorkbookName = safeGetWorkbookName ?? throw new ArgumentNullException(nameof(safeGetWorkbookName));
            _onCasePaneUpdatedNotification = onCasePaneUpdatedNotification;
        }

        internal void NotifyCasePaneUpdatedIfNeeded(
            Excel.Workbook workbook,
            string reason,
            TaskPaneSnapshotBuilderService.TaskPaneBuildResult buildResult,
            bool? originalSavedState = null)
        {
            if (workbook == null)
            {
                return;
            }

            try
            {
                bool updatedCaseSnapshotCache = buildResult != null && buildResult.UpdatedCaseSnapshotCache;
                if (updatedCaseSnapshotCache)
                {
                    RestoreWorkbookSavedState(workbook, originalSavedState);
                }

                if (!CasePaneCacheRefreshNotificationPolicy.ShouldNotify(updatedCaseSnapshotCache, reason))
                {
                    return;
                }

                if (_onCasePaneUpdatedNotification != null)
                {
                    _onCasePaneUpdatedNotification(reason ?? string.Empty);
                    return;
                }

                MessageBox.Show("文書ボタンパネルを更新しました", ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                _logger.Info("CASE pane cache refresh notification was shown. workbook=" + _safeGetWorkbookName(workbook) + ", reason=" + (reason ?? string.Empty));
            }
            catch (Exception ex)
            {
                _logger.Error("NotifyCasePaneUpdatedIfNeeded failed.", ex);
            }
        }

        internal bool? TryGetWorkbookSavedState(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            try
            {
                return workbook.Saved;
            }
            catch (Exception ex)
            {
                _logger.Error("TryGetWorkbookSavedState failed.", ex);
                return null;
            }
        }

        internal void RestoreWorkbookSavedState(Excel.Workbook workbook, bool? originalSavedState)
        {
            if (workbook == null || !originalSavedState.HasValue)
            {
                return;
            }

            try
            {
                workbook.Saved = originalSavedState.Value;
            }
            catch (Exception ex)
            {
                _logger.Error("RestoreWorkbookSavedState failed.", ex);
            }
        }
    }
}
