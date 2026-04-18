using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal static class TaskPaneSnapshotFormat
    {
        internal const string ExportVersion = "2";
        private const string MetaPrefix = "META\t";

        /// <summary>
        internal static bool IsCompatible(string snapshotText)
        {
            string version = TryReadExportVersion(snapshotText);
            return string.Equals(version, ExportVersion, StringComparison.Ordinal);
        }

        /// <summary>
        internal static string TryReadExportVersion(string snapshotText)
        {
            if (string.IsNullOrWhiteSpace(snapshotText))
            {
                return string.Empty;
            }

            string normalized = snapshotText.Replace("\r\n", "\n");
            string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0)
            {
                return string.Empty;
            }

            string metaLine = lines[0] ?? string.Empty;
            if (!metaLine.StartsWith(MetaPrefix, StringComparison.Ordinal))
            {
                return string.Empty;
            }

            string[] fields = metaLine.Split('\t');
            return fields.Length > 1 ? fields[1] ?? string.Empty : string.Empty;
        }
    }
}
