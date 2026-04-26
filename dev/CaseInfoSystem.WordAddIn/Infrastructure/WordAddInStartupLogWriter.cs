using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace CaseInfoSystem.WordAddIn.Infrastructure
{
    internal static class WordAddInStartupLogWriter
    {
        private const string LogDirectoryName = "CaseInfoSystem.WordAddIn";
        private const string LogFileName = "CaseInfoSystem.WordAddIn_startup.log";
        private const string FallbackMarker = " [primaryWriteFailed][tempFallback]";
        private static readonly Encoding LogEncoding = new UTF8Encoding(false);

        internal static string GetPrimaryLogPath()
        {
            string localAppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppDataPath, LogDirectoryName, "logs", LogFileName);
        }

        internal static string GetFallbackLogPath()
        {
            return Path.Combine(Path.GetTempPath(), LogDirectoryName, LogFileName);
        }

        internal static void Write(string message)
        {
            string line = BuildHeader() + (message ?? string.Empty);
            if (TryAppend(GetPrimaryLogPath(), line))
            {
                return;
            }

            TryAppend(GetFallbackLogPath(), line + FallbackMarker);
        }

        internal static void WriteException(string context, Exception exception)
        {
            var builder = new StringBuilder();
            builder.AppendLine(context ?? "Unhandled exception");

            if (exception == null)
            {
                builder.Append("Exception=<null>");
                Write(builder.ToString());
                return;
            }

            int depth = 0;
            Exception current = exception;
            while (current != null)
            {
                builder.Append("Exception[").Append(depth.ToString(CultureInfo.InvariantCulture)).Append("].Type=")
                    .AppendLine(current.GetType().FullName ?? string.Empty);
                builder.Append("Exception[").Append(depth.ToString(CultureInfo.InvariantCulture)).Append("].Message=")
                    .AppendLine(current.Message ?? string.Empty);
                builder.Append("Exception[").Append(depth.ToString(CultureInfo.InvariantCulture)).Append("].StackTrace=")
                    .AppendLine(current.StackTrace ?? string.Empty);

                current = current.InnerException;
                depth++;
            }

            Write(builder.ToString().TrimEnd());
        }

        private static string BuildHeader()
        {
            Process currentProcess = Process.GetCurrentProcess();
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
                + " [ProcessName=" + currentProcess.ProcessName + "]"
                + " [PID=" + currentProcess.Id.ToString(CultureInfo.InvariantCulture) + "] ";
        }

        private static bool TryAppend(string path, string content)
        {
            try
            {
                string directoryPath = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                File.AppendAllText(path, content + Environment.NewLine, LogEncoding);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
