using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class ExcelAddInTraceLogWriter
    {
        private const string SystemRootFolderName = "案件情報System";
        private const string TraceLogFileName = "CaseInfoSystem.ExcelAddIn_trace.log";
        private const string FallbackMarker = " [primaryWriteFailed][tempFallback]";
        private static readonly Encoding TraceLogEncoding = new UTF8Encoding(false);

        internal static string GetPrimaryTraceLogRelativePath()
        {
            return Path.Combine("logs", TraceLogFileName);
        }

        internal static string GetPrimaryTraceLogPath()
        {
            return Path.Combine(ResolveLogDirectory(), TraceLogFileName);
        }

        internal static void Write(string message)
        {
            string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
                + " [PID=" + Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture) + "] CaseInfoSystem: "
                + (message ?? string.Empty);

            if (TryAppendToPrimaryLog(line))
            {
                return;
            }

            TryAppendToFallbackLog(line);
        }

        private static string ResolveLogDirectory()
        {
            return Path.Combine(GetPrimarySystemRootPath(), "logs");
        }

        internal static string GetPrimarySystemRootPath()
        {
            string localDocuments = GetLocalDocumentsPath();
            string redirectedDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return ResolvePrimarySystemRootPath(localDocuments, redirectedDocuments);
        }

        internal static string ResolvePrimarySystemRootPath(string localDocumentsPath, string redirectedDocumentsPath)
        {
            string existingSystemRootPath = TryResolveExistingSystemRootPath(localDocumentsPath, redirectedDocumentsPath);
            if (!string.IsNullOrWhiteSpace(existingSystemRootPath))
            {
                return existingSystemRootPath;
            }

            return BuildSystemRootPath(redirectedDocumentsPath);
        }

        private static string TryResolveExistingSystemRootPath(params string[] documentRoots)
        {
            if (documentRoots == null)
            {
                return string.Empty;
            }

            for (int i = 0; i < documentRoots.Length; i++)
            {
                string systemRootPath = BuildSystemRootPath(documentRoots[i]);
                if (LooksLikeSystemRoot(systemRootPath))
                {
                    return systemRootPath;
                }
            }

            return string.Empty;
        }

        private static string GetLocalDocumentsPath()
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrWhiteSpace(userProfile))
            {
                return string.Empty;
            }

            return Path.Combine(userProfile, "Documents");
        }

        private static string BuildSystemRootPath(string documentsPath)
        {
            if (string.IsNullOrWhiteSpace(documentsPath))
            {
                return string.Empty;
            }

            return Path.Combine(documentsPath, SystemRootFolderName);
        }

        private static bool LooksLikeSystemRoot(string systemRootPath)
        {
            if (string.IsNullOrWhiteSpace(systemRootPath) || !Directory.Exists(systemRootPath))
            {
                return false;
            }

            if (Directory.Exists(Path.Combine(systemRootPath, "Addins", "CaseInfoSystem.ExcelAddIn")))
            {
                return true;
            }

            return File.Exists(Path.Combine(systemRootPath, "案件情報System_Kernel.xlsx"));
        }

        private static bool TryAppendToPrimaryLog(string line)
        {
            try
            {
                AppendLine(GetPrimaryTraceLogPath(), line);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void TryAppendToFallbackLog(string line)
        {
            try
            {
                AppendLine(Path.Combine(Path.GetTempPath(), "CaseInfoSystem.ExcelAddIn", TraceLogFileName), line + FallbackMarker);
            }
            catch
            {
            }
        }

        private static void AppendLine(string logPath, string line)
        {
            string logDirectory = Path.GetDirectoryName(logPath) ?? ResolveLogDirectory();
            Directory.CreateDirectory(logDirectory);
            File.AppendAllText(logPath, line + Environment.NewLine, TraceLogEncoding);
        }
    }
}
