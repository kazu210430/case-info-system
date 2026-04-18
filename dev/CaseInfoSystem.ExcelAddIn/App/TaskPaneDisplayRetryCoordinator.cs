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
            Action<Excel.Workbook, string, int> waitBeforeRetry,
            Action onShown,
            Action<Excel.Workbook, string> scheduleFallback)
        {
            for (int attempt = 0; attempt < _maxAttempts; attempt++)
            {
                int attemptNumber = attempt + 1;
                if (tryShowOnce(workbook, reason, attemptNumber))
                {
                    onShown();
                    return;
                }

                waitBeforeRetry(workbook, reason, attemptNumber);
            }

            scheduleFallback(workbook, reason);
        }
    }
}
