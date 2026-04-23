using System;
using System.Diagnostics;
using System.Globalization;
using System.Management;
using System.Text;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class ExcelProcessLaunchContextTracer
    {
        private const string DetailedStartupDiagnosticsEnvironmentVariableName = "CASEINFOSYSTEM_EXCEL_STARTUP_TRACE";

        internal static void Trace(Logger logger)
        {
            if (logger == null)
            {
                return;
            }

            try
            {
                using (Process current = Process.GetCurrentProcess())
                {
                    int currentProcessId = current.Id;
                    bool collectDetailedDiagnostics = IsDetailedStartupDiagnosticsEnabled();
                    var messageBuilder = new StringBuilder();
                    messageBuilder.Append("Process launch context. currentPid=");
                    messageBuilder.Append(currentProcessId.ToString(CultureInfo.InvariantCulture));
                    messageBuilder.Append(", currentName=");
                    messageBuilder.Append(SafeProcessName(current));
                    messageBuilder.Append(", sessionId=");
                    messageBuilder.Append(current.SessionId.ToString(CultureInfo.InvariantCulture));
                    messageBuilder.Append(", startTime=");
                    messageBuilder.Append(SafeProcessStartTime(current));
                    messageBuilder.Append(", startupDiagnosticsMode=");
                    messageBuilder.Append(collectDetailedDiagnostics ? "detailed" : "basic");

                    if (collectDetailedDiagnostics)
                    {
                        int parentProcessId = TryGetParentProcessId(currentProcessId);
                        messageBuilder.Append(", parentPid=");
                        messageBuilder.Append(parentProcessId.ToString(CultureInfo.InvariantCulture));
                        messageBuilder.Append(", parent=");
                        messageBuilder.Append(GetProcessSummary(parentProcessId));
                        messageBuilder.Append(", commandLine=");
                        messageBuilder.Append(SafeCurrentCommandLine());
                        messageBuilder.Append(", excelProcesses=[");
                        messageBuilder.Append(BuildExcelProcessSnapshot(currentProcessId));
                        messageBuilder.Append("]");
                    }

                    logger.Info(messageBuilder.ToString());
                }
            }
            catch (Exception ex)
            {
                logger.Error("TraceProcessLaunchContext failed.", ex);
            }
        }

        private static bool IsDetailedStartupDiagnosticsEnabled()
        {
            string value = Environment.GetEnvironmentVariable(DetailedStartupDiagnosticsEnvironmentVariableName);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string normalizedValue = value.Trim();
            return normalizedValue.Equals("1", StringComparison.OrdinalIgnoreCase)
                || normalizedValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                || normalizedValue.Equals("on", StringComparison.OrdinalIgnoreCase)
                || normalizedValue.Equals("full", StringComparison.OrdinalIgnoreCase)
                || normalizedValue.Equals("detailed", StringComparison.OrdinalIgnoreCase);
        }

        private static int TryGetParentProcessId(int processId)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = " + processId.ToString(CultureInfo.InvariantCulture)))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {
                        object parentProcessId = process["ParentProcessId"];
                        if (parentProcessId == null)
                        {
                            return 0;
                        }

                        return Convert.ToInt32(parentProcessId, CultureInfo.InvariantCulture);
                    }
                }
            }
            catch
            {
                return 0;
            }

            return 0;
        }

        private static string TryGetProcessCommandLine(int processId)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + processId.ToString(CultureInfo.InvariantCulture)))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {
                        object commandLine = process["CommandLine"];
                        return commandLine == null ? string.Empty : Convert.ToString(commandLine, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        private static string SafeCurrentCommandLine()
        {
            try
            {
                return Environment.CommandLine ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetProcessSummary(int processId)
        {
            if (processId <= 0)
            {
                return "(unknown)";
            }

            try
            {
                using (Process process = Process.GetProcessById(processId))
                {
                    return "pid="
                        + process.Id.ToString(CultureInfo.InvariantCulture)
                        + ",name="
                        + SafeProcessName(process)
                        + ",startTime="
                        + SafeProcessStartTime(process);
                }
            }
            catch
            {
                return "pid=" + processId.ToString(CultureInfo.InvariantCulture) + ",name=(unavailable)";
            }
        }

        private static string BuildExcelProcessSnapshot(int currentProcessId)
        {
            var builder = new StringBuilder();

            try
            {
                Process[] processes = Process.GetProcessesByName("EXCEL");
                foreach (Process process in processes)
                {
                    using (process)
                    {
                        if (builder.Length > 0)
                        {
                            builder.Append(" | ");
                        }

                        builder.Append("pid=");
                        builder.Append(process.Id.ToString(CultureInfo.InvariantCulture));
                        builder.Append(",name=");
                        builder.Append(SafeProcessName(process));
                        builder.Append(",startTime=");
                        builder.Append(SafeProcessStartTime(process));
                        builder.Append(",isCurrent=");
                        builder.Append((process.Id == currentProcessId).ToString());
                        builder.Append(",commandLine=");
                        builder.Append(TryGetProcessCommandLine(process.Id));
                    }
                }
            }
            catch
            {
                return builder.ToString();
            }

            return builder.ToString();
        }

        private static string SafeProcessName(Process process)
        {
            try
            {
                return process == null ? string.Empty : process.ProcessName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeProcessStartTime(Process process)
        {
            try
            {
                if (process == null)
                {
                    return string.Empty;
                }

                return process.StartTime.ToString("O", CultureInfo.InvariantCulture);
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
