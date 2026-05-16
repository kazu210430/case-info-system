using System;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class AddInExecutionBoundaryCoordinator : IScreenUpdatingExecutionBridge, ITaskPaneRefreshSuppressionBridge
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Excel.Application _application;
        private readonly Logger _logger;
        private readonly Func<string> _formatActiveExcelState;
        private int _taskPaneRefreshSuppressionCount;

        internal AddInExecutionBoundaryCoordinator(
            Excel.Application application,
            Logger logger,
            Func<string> formatActiveExcelState)
        {
            _application = application;
            _logger = logger;
            _formatActiveExcelState = formatActiveExcelState;
        }

        internal int TaskPaneRefreshSuppressionCount
        {
            get { return _taskPaneRefreshSuppressionCount; }
        }

        public void Execute(Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            bool previousScreenUpdating = true;
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                action();
            }
            finally
            {
                try
                {
                    _application.ScreenUpdating = previousScreenUpdating;
                }
                catch
                {
                    // ScreenUpdating restore failure must not mask the completed business action.
                }
            }
        }

        public IDisposable Enter(string reason)
        {
            _taskPaneRefreshSuppressionCount++;
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=suppress-enter reason="
                + (reason ?? string.Empty)
                + ", suppressionCount="
                + _taskPaneRefreshSuppressionCount.ToString(CultureInfo.InvariantCulture)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "Task pane refresh suppression entered. reason="
                + (reason ?? string.Empty)
                + ", suppressionCount="
                + _taskPaneRefreshSuppressionCount.ToString());

            return new DelegateDisposable(() => ExitSuppression(reason));
        }

        private void ExitSuppression(string reason)
        {
            if (_taskPaneRefreshSuppressionCount > 0)
            {
                _taskPaneRefreshSuppressionCount--;
            }

            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=ThisAddIn action=suppress-exit reason="
                + (reason ?? string.Empty)
                + ", suppressionCount="
                + _taskPaneRefreshSuppressionCount.ToString(CultureInfo.InvariantCulture)
                + ", activeState="
                + FormatActiveExcelState());
            _logger?.Info(
                "Task pane refresh suppression exited. reason="
                + (reason ?? string.Empty)
                + ", suppressionCount="
                + _taskPaneRefreshSuppressionCount.ToString());
        }

        private string FormatActiveExcelState()
        {
            try
            {
                return _formatActiveExcelState == null ? string.Empty : _formatActiveExcelState();
            }
            catch
            {
                return string.Empty;
            }
        }

        private sealed class DelegateDisposable : IDisposable
        {
            private readonly Action _disposeAction;
            private bool _disposed;

            internal DelegateDisposable(Action disposeAction)
            {
                _disposeAction = disposeAction ?? throw new ArgumentNullException(nameof(disposeAction));
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _disposeAction();
            }
        }
    }
}
