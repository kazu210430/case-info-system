using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // Orchestrates replace/register/remove over the shared host map.
    // The shared map itself is still owned by TaskPaneManager, concrete VSTO pane lifetime stays below TaskPaneHost/ThisAddIn,
    // TaskPaneHostFactory composition now lives outside this type, and the host descriptor formatter is consumed only for diagnostics.
    internal sealed class TaskPaneHostRegistry
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";

        private readonly Dictionary<string, TaskPaneHost> _hostsByWindowKey;
        private readonly Logger _logger;
        private readonly Func<TaskPaneHost, string> _formatHostDescriptorForDiagnostics;
        private readonly TaskPaneHostFactory _taskPaneHostFactory;

        internal TaskPaneHostRegistry(
            Dictionary<string, TaskPaneHost> hostsByWindowKey,
            Logger logger,
            Func<TaskPaneHost, string> formatHostDescriptorForDiagnostics,
            TaskPaneHostFactory taskPaneHostFactory)
        {
            _hostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            // Diagnostic-only input: registry does not own host identity, metadata timing, or replace/remove decisions through this formatter.
            _formatHostDescriptorForDiagnostics = formatHostDescriptorForDiagnostics ?? throw new ArgumentNullException(nameof(formatHostDescriptorForDiagnostics));
            _taskPaneHostFactory = taskPaneHostFactory ?? throw new ArgumentNullException(nameof(taskPaneHostFactory));
        }

        internal void RegisterHost(TaskPaneHost host)
        {
            if (host == null)
            {
                throw new ArgumentNullException(nameof(host));
            }

            if (TryGetDifferentRegisteredHost(host, out TaskPaneHost existingHost))
            {
                DisposeHostForReplacement(existingHost);
            }

            StoreRegisteredHost(host);
        }

        internal TaskPaneHost GetOrReplaceHost(string windowKey, Excel.Window window, WorkbookRole role)
        {
            if (TryGetReusableHost(windowKey, role, out TaskPaneHost existingHost))
            {
                return existingHost;
            }

            RemoveExistingHostForReplacement(windowKey);
            return CreateAndRegisterHost(windowKey, window, role);
        }

        internal void RemoveWorkbookPanes(string workbookFullName)
        {
            foreach (string windowKey in CollectWindowKeysForWorkbook(workbookFullName))
            {
                RemoveHost(windowKey);
            }
        }

        internal void DisposeAll()
        {
            foreach (TaskPaneHost host in SnapshotHosts())
            {
                DisposeHostForReplacement(host);
            }

            _hostsByWindowKey.Clear();
        }

        internal void RemoveHost(string windowKey)
        {
            if (string.IsNullOrWhiteSpace(windowKey))
            {
                return;
            }

            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                return;
            }

            LogHostRemoval(host);
            _hostsByWindowKey.Remove(windowKey);
            DisposeHostAfterRemoval(windowKey, host);
        }

        private bool TryGetDifferentRegisteredHost(TaskPaneHost host, out TaskPaneHost existingHost)
        {
            if (_hostsByWindowKey.TryGetValue(host.WindowKey, out existingHost))
            {
                return !ReferenceEquals(existingHost, host);
            }

            return false;
        }

        private void StoreRegisteredHost(TaskPaneHost host)
        {
            _hostsByWindowKey[host.WindowKey] = host;
        }

        private bool TryGetReusableHost(string windowKey, WorkbookRole role, out TaskPaneHost host)
        {
            if (_hostsByWindowKey.TryGetValue(windowKey, out host))
            {
                return IsHostCompatibleWithRole(host, role);
            }

            return false;
        }

        private static bool IsHostCompatibleWithRole(TaskPaneHost host, WorkbookRole role)
        {
            return (role == WorkbookRole.Kernel && host.Control is KernelNavigationControl)
                || (role == WorkbookRole.Case && host.Control is DocumentButtonsControl)
                || (role == WorkbookRole.Accounting && host.Control is AccountingNavigationControl);
        }

        private void RemoveExistingHostForReplacement(string windowKey)
        {
            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost existingHost))
            {
                return;
            }

            DisposeHostForReplacement(existingHost);
            _hostsByWindowKey.Remove(windowKey);
        }

        private TaskPaneHost CreateAndRegisterHost(string windowKey, Excel.Window window, WorkbookRole role)
        {
            // Factory owns control creation and ActionInvoked binding. Registry only decides reuse vs replace and records the host.
            TaskPaneHost host = _taskPaneHostFactory.CreateHost(windowKey, window, role, out string paneRoleName);
            _hostsByWindowKey.Add(windowKey, host);
            _logger.Debug(nameof(TaskPaneHostRegistry), "TaskPane host registered. role=" + paneRoleName + ", windowKey=" + windowKey);
            return host;
        }

        private List<string> CollectWindowKeysForWorkbook(string workbookFullName)
        {
            var targetKeys = new List<string>();
            foreach (KeyValuePair<string, TaskPaneHost> pair in _hostsByWindowKey)
            {
                if (string.Equals(pair.Value.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                {
                    targetKeys.Add(pair.Key);
                }
            }

            return targetKeys;
        }

        private List<TaskPaneHost> SnapshotHosts()
        {
            return new List<TaskPaneHost>(_hostsByWindowKey.Values);
        }

        private static void DisposeHostForReplacement(TaskPaneHost host)
        {
            host.Dispose();
        }

        private void DisposeHostAfterRemoval(string windowKey, TaskPaneHost host)
        {
            try
            {
                host.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Error("TaskPane host dispose failed. windowKey=" + windowKey, ex);
            }
        }

        private void LogHostRemoval(TaskPaneHost host)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=remove-host host="
                + _formatHostDescriptorForDiagnostics(host));
        }
    }
}
