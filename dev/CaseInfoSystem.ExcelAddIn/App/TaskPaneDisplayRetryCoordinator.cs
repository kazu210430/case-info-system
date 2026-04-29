using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneDisplayRetryCoordinator
    {
        private readonly int _maxAttempts;

        internal TaskPaneDisplayRetryCoordinator(int maxAttempts)
        {
            _maxAttempts = maxAttempts;
        }

        internal void ShowWhenReady(
            Excel.Workbook workbook,
            string reason,
            Func<Excel.Workbook, string, int, bool> tryShowOnce,
            Action<Excel.Workbook, string, int, Action> scheduleRetry,
            Action onShown,
            Action<Excel.Workbook, string> scheduleFallback)
        {
            if (tryShowOnce(workbook, reason, 1))
            {
                onShown();
                return;
            }

            ScheduleRetryAttempt(2);

            void ScheduleRetryAttempt(int attemptNumber)
            {
                if (attemptNumber > _maxAttempts)
                {
                    scheduleFallback(workbook, reason);
                    return;
                }

                scheduleRetry(
                    workbook,
                    reason,
                    attemptNumber,
                    () =>
                    {
                        if (tryShowOnce(workbook, reason, attemptNumber))
                        {
                            onShown();
                            return;
                        }

                        ScheduleRetryAttempt(attemptNumber + 1);
                    });
            }
        }
    }
}
