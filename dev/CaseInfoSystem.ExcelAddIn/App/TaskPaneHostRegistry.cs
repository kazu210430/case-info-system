using System;
using System.Collections.Generic;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // Remove/replace/register orchestration owner over the shared host map.
    // The shared map itself is still owned by TaskPaneManager, concrete VSTO pane lifetime stays below TaskPaneHost/ThisAddIn,
    // TaskPaneHostFactory composition now lives outside this type, and the host descriptor formatter is consumed only for diagnostics.
    // WorkbookFullName is used here only as a remove-selection input, not as metadata ownership.
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
                DisposeHostWithoutRegistryMutation(existingHost, "replace-different-registered-host");
            }

            StoreRegisteredHost(host);
        }

        internal TaskPaneHost GetOrReplaceHost(string windowKey, Excel.Window window, WorkbookRole role)
        {
            // Registry is the shared-map orchestration owner here:
            // reuse if compatible, otherwise dispose+unregister for replacement, then create+register a new host.
            if (TryGetReusableRegisteredHost(windowKey, role, out TaskPaneHost existingHost))
            {
                return existingHost;
            }

            DisposeThenUnregisterHostForReplacement(windowKey);
            return CreateHostThenRegister(windowKey, window, role);
        }

        internal void RemoveWorkbookPanes(string workbookFullName)
        {
            // WorkbookFullName is consumed only to choose which registered window keys should be removed.
            foreach (string windowKey in CollectWindowKeysForWorkbookRemovalSelection(workbookFullName))
            {
                RemoveHost(windowKey);
            }
        }

        internal void DisposeAll()
        {
            // Shutdown cleanup fixed point:
            // snapshot registered hosts -> dispose each host -> clear the shared map.
            List<TaskPaneHost> hosts = SnapshotHosts();
            string roleCounts = FormatHostRoleCounts(hosts);
            _logger.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=dispose-all-start hostCount="
                + hosts.Count.ToString(CultureInfo.InvariantCulture)
                + ", "
                + roleCounts);
            foreach (TaskPaneHost host in hosts)
            {
                DisposeHostWithoutRegistryMutation(host, "dispose-all");
            }

            _hostsByWindowKey.Clear();
            _logger.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=dispose-all-complete disposedCount="
                + hosts.Count.ToString(CultureInfo.InvariantCulture)
                + ", "
                + roleCounts);
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

            // Standard remove fixed point:
            // log while the host is still registered -> remove the shared-map entry -> dispose the now-unregistered host.
            LogHostRemoval(host);
            _hostsByWindowKey.Remove(windowKey);
            bool disposeSucceeded = DisposeHostAfterRegistryRemoval(windowKey, host);
            LogHostRemovalComplete(windowKey, host, disposeSucceeded);
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

        private bool TryGetReusableRegisteredHost(string windowKey, WorkbookRole role, out TaskPaneHost host)
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

        private void DisposeThenUnregisterHostForReplacement(string windowKey)
        {
            if (!_hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost existingHost))
            {
                return;
            }

            // Replacement remove ordering is intentionally fixed in current-state:
            // dispose first, then drop the shared-map entry.
            DisposeHostWithoutRegistryMutation(existingHost, "replace-registered-host");
            _hostsByWindowKey.Remove(windowKey);
        }

        private TaskPaneHost CreateHostThenRegister(string windowKey, Excel.Window window, WorkbookRole role)
        {
            // Factory owns control creation and ActionInvoked binding.
            // Registry owns only the reuse/replace/register orchestration over the shared map.
            TaskPaneHost host = _taskPaneHostFactory.CreateHost(windowKey, window, role, out string paneRoleName);
            _hostsByWindowKey.Add(windowKey, host);
            _logger.Debug(nameof(TaskPaneHostRegistry), "TaskPane host registered. role=" + paneRoleName + ", windowKey=" + windowKey);
            return host;
        }

        private List<string> CollectWindowKeysForWorkbookRemovalSelection(string workbookFullName)
        {
            // WorkbookFullName is consumed only for workbook-scope remove selection here.
            // Registry does not own metadata write timing; it only reads the already-populated value.
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

        private void DisposeHostWithoutRegistryMutation(TaskPaneHost host, string context)
        {
            string windowKey = host?.WindowKey ?? string.Empty;
            LogHostDisposeStart(context, windowKey, host);
            try
            {
                host.Dispose();
                LogHostDisposeComplete(context, windowKey, host);
            }
            catch (Exception ex)
            {
                LogHostDisposeFailure(context, windowKey, host, ex);
                throw;
            }
        }

        private bool DisposeHostAfterRegistryRemoval(string windowKey, TaskPaneHost host)
        {
            LogHostDisposeStart("remove-host", windowKey, host);
            try
            {
                host.Dispose();
                LogHostDisposeComplete("remove-host", windowKey, host);
                return true;
            }
            catch (Exception ex)
            {
                LogHostDisposeFailure("remove-host", windowKey, host, ex);
                _logger.Error("TaskPane host dispose failed. windowKey=" + windowKey, ex);
                return false;
            }
        }

        private void LogHostRemoval(TaskPaneHost host)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=remove-host host="
                + _formatHostDescriptorForDiagnostics(host));
        }

        private void LogHostRemovalComplete(string windowKey, TaskPaneHost host, bool disposeSucceeded)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=remove-host-complete disposeSucceeded="
                + disposeSucceeded.ToString()
                + ", windowKey="
                + (windowKey ?? string.Empty)
                + ", host="
                + FormatSafeHostDescriptor(host));
        }

        private void LogHostDisposeStart(string context, string windowKey, TaskPaneHost host)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=host-dispose-start context="
                + (context ?? string.Empty)
                + ", windowKey="
                + (windowKey ?? string.Empty)
                + ", host="
                + FormatSafeHostDescriptor(host));
        }

        private void LogHostDisposeComplete(string context, string windowKey, TaskPaneHost host)
        {
            _logger?.Info(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=host-dispose-complete context="
                + (context ?? string.Empty)
                + ", windowKey="
                + (windowKey ?? string.Empty)
                + ", host="
                + FormatSafeHostDescriptor(host));
        }

        private void LogHostDisposeFailure(string context, string windowKey, TaskPaneHost host, Exception exception)
        {
            _logger?.Error(
                KernelFlickerTracePrefix
                + " source=TaskPaneManager action=host-dispose-failure context="
                + (context ?? string.Empty)
                + ", windowKey="
                + (windowKey ?? string.Empty)
                + ", host="
                + FormatSafeHostDescriptor(host)
                + ", exceptionType="
                + exception.GetType().Name
                + ", hResult=0x"
                + exception.HResult.ToString("X8", CultureInfo.InvariantCulture)
                + ", message="
                + exception.Message,
                exception);
        }

        private static string FormatSafeHostDescriptor(TaskPaneHost host)
        {
            if (host == null)
            {
                return "paneRole=Unknown, windowKey=, workbookFullName=, controlType=";
            }

            string controlType = host.Control?.GetType().Name ?? string.Empty;
            return "paneRole="
                + GetSafePaneRoleName(host)
                + ", windowKey="
                + (host.WindowKey ?? string.Empty)
                + ", workbookFullName="
                + (host.WorkbookFullName ?? string.Empty)
                + ", controlType="
                + controlType;
        }

        private static string FormatHostRoleCounts(IEnumerable<TaskPaneHost> hosts)
        {
            int kernelHostCount = 0;
            int caseHostCount = 0;
            int accountingHostCount = 0;
            int unknownHostCount = 0;

            foreach (TaskPaneHost host in hosts)
            {
                string roleName = GetSafePaneRoleName(host);
                if (string.Equals(roleName, "Kernel", StringComparison.OrdinalIgnoreCase))
                {
                    kernelHostCount++;
                    continue;
                }

                if (string.Equals(roleName, "Case", StringComparison.OrdinalIgnoreCase))
                {
                    caseHostCount++;
                    continue;
                }

                if (string.Equals(roleName, "Accounting", StringComparison.OrdinalIgnoreCase))
                {
                    accountingHostCount++;
                    continue;
                }

                unknownHostCount++;
            }

            return "kernelHostCount="
                + kernelHostCount.ToString(CultureInfo.InvariantCulture)
                + ", caseHostCount="
                + caseHostCount.ToString(CultureInfo.InvariantCulture)
                + ", accountingHostCount="
                + accountingHostCount.ToString(CultureInfo.InvariantCulture)
                + ", unknownHostCount="
                + unknownHostCount.ToString(CultureInfo.InvariantCulture);
        }

        private static string GetSafePaneRoleName(TaskPaneHost host)
        {
            if (host?.Control is KernelNavigationControl)
            {
                return "Kernel";
            }

            if (host?.Control is DocumentButtonsControl)
            {
                return "Case";
            }

            if (host?.Control is AccountingNavigationControl)
            {
                return "Accounting";
            }

            return "Unknown";
        }
    }
}
