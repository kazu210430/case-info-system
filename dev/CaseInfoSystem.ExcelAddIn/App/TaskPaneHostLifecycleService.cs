using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneHostLifecycleService
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly IDictionary<string, TaskPaneHost> _hostsByWindowKey;
        private readonly TaskPaneHostRegistry _taskPaneHostRegistry;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;
        private readonly Func<TaskPaneHost, string> _formatHostDescriptor;

        internal TaskPaneHostLifecycleService(
            IDictionary<string, TaskPaneHost> hostsByWindowKey,
            TaskPaneHostRegistry taskPaneHostRegistry,
            ExcelInteropService excelInteropService,
            Logger logger,
            Func<TaskPaneHost, string> formatHostDescriptor)
        {
            _hostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
            _taskPaneHostRegistry = taskPaneHostRegistry ?? throw new ArgumentNullException(nameof(taskPaneHostRegistry));
            _excelInteropService = excelInteropService;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _formatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
        }

        internal TaskPaneHost ResolveRefreshHost(WorkbookContext context, string windowKey, int refreshPaneCallId)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            RemoveStaleKernelHosts(context, windowKey);
            TaskPaneHost host = _taskPaneHostRegistry.GetOrReplaceHost(windowKey, context.Window, context.Role);
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=host-selected refreshPaneCallId="
                + refreshPaneCallId.ToString()
                + ", host="
                + _formatHostDescriptor(host));
            return host;
        }

        internal void RegisterHost(TaskPaneHost host)
        {
            _taskPaneHostRegistry.RegisterHost(host);
        }

        internal void RemoveWorkbookPanes(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            string workbookFullName = GetWorkbookFullName(workbook);
            if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                return;
            }

            _taskPaneHostRegistry.RemoveWorkbookPanes(workbookFullName);
        }

        internal void RemoveHost(string windowKey)
        {
            _taskPaneHostRegistry.RemoveHost(windowKey);
        }

        internal void DisposeAll()
        {
            _taskPaneHostRegistry.DisposeAll();
        }

        private void RemoveStaleKernelHosts(WorkbookContext context, string activeWindowKey)
        {
            if (context == null
                || context.Role != WorkbookRole.Kernel
                || string.IsNullOrWhiteSpace(context.WorkbookFullName)
                || string.IsNullOrWhiteSpace(activeWindowKey))
            {
                return;
            }

            var staleKeys = new List<string>();
            foreach (KeyValuePair<string, TaskPaneHost> pair in _hostsByWindowKey)
            {
                if (string.Equals(pair.Key, activeWindowKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                TaskPaneHost host = pair.Value;
                if (!IsKernelHost(host))
                {
                    continue;
                }

                if (!string.Equals(host.WorkbookFullName, context.WorkbookFullName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                staleKeys.Add(pair.Key);
            }

            foreach (string staleKey in staleKeys)
            {
                _logger.Info(
                    "Removed stale kernel task pane host. workbook="
                    + context.WorkbookFullName
                    + ", staleWindowKey="
                    + staleKey
                    + ", activeWindowKey="
                    + activeWindowKey);
                _taskPaneHostRegistry.RemoveHost(staleKey);
            }
        }

        private string GetWorkbookFullName(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            if (_excelInteropService != null)
            {
                return _excelInteropService.GetWorkbookFullName(workbook) ?? string.Empty;
            }

            return workbook.FullName ?? string.Empty;
        }

        private static bool IsKernelHost(TaskPaneHost host)
        {
            return host != null && host.Control is KernelNavigationControl;
        }
    }
}
