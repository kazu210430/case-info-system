using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class ManagedWorkbookCloseMarkerStore
    {
        internal const int DefaultTimeToLiveSeconds = 15;
        private const string MarkerFileName = "managed-workbook-close.marker";
        private const string VersionValue = "1";
        private readonly string _markerPath;
        private readonly Func<DateTime> _utcNow;

        internal ManagedWorkbookCloseMarkerStore()
            : this(null, null)
        {
        }

        internal ManagedWorkbookCloseMarkerStore(string markerPath, Func<DateTime> utcNow = null)
        {
            _markerPath = string.IsNullOrWhiteSpace(markerPath) ? ResolveDefaultMarkerPath() : markerPath;
            _utcNow = utcNow ?? (() => DateTime.UtcNow);
        }

        internal string MarkerPath
        {
            get { return _markerPath; }
        }

        internal void Write(ManagedWorkbookCloseMarkerKind kind, string workbookKey)
        {
            string directory = Path.GetDirectoryName(_markerPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            DateTime createdUtc = _utcNow();
            string[] lines =
            {
                "version=" + VersionValue,
                "kind=" + kind.ToString(),
                "createdUtc=" + createdUtc.ToString("O", CultureInfo.InvariantCulture),
                "ttlSeconds=" + DefaultTimeToLiveSeconds.ToString(CultureInfo.InvariantCulture),
                "workbookKeyBase64=" + Encode(workbookKey)
            };

            File.WriteAllLines(_markerPath, lines, new UTF8Encoding(false));
        }

        internal ManagedWorkbookCloseMarkerReadResult Consume()
        {
            if (!File.Exists(_markerPath))
            {
                return ManagedWorkbookCloseMarkerReadResult.NoMarker(_markerPath);
            }

            ManagedWorkbookCloseMarker marker;
            try
            {
                marker = Parse(File.ReadAllLines(_markerPath, Encoding.UTF8));
            }
            catch
            {
                TryDelete();
                return ManagedWorkbookCloseMarkerReadResult.Invalid(_markerPath);
            }

            if (marker == null)
            {
                TryDelete();
                return ManagedWorkbookCloseMarkerReadResult.Invalid(_markerPath);
            }

            TimeSpan age = _utcNow() - marker.CreatedUtc;
            if (age < TimeSpan.Zero)
            {
                age = TimeSpan.Zero;
            }

            if (age > TimeSpan.FromSeconds(marker.TimeToLiveSeconds))
            {
                TryDelete();
                return ManagedWorkbookCloseMarkerReadResult.Expired(_markerPath, marker, age);
            }

            TryDelete();
            return ManagedWorkbookCloseMarkerReadResult.Valid(_markerPath, marker, age);
        }

        private static string ResolveDefaultMarkerPath()
        {
            return Path.Combine(ExcelAddInTraceLogWriter.GetPrimarySystemRootPath(), "logs", MarkerFileName);
        }

        private ManagedWorkbookCloseMarker Parse(IEnumerable<string> lines)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                int separatorIndex = line.IndexOf('=');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                values[line.Substring(0, separatorIndex)] = line.Substring(separatorIndex + 1);
            }

            string version;
            string kindText;
            string createdUtcText;
            string ttlSecondsText;
            if (!values.TryGetValue("version", out version)
                || !string.Equals(version, VersionValue, StringComparison.Ordinal)
                || !values.TryGetValue("kind", out kindText)
                || !values.TryGetValue("createdUtc", out createdUtcText)
                || !values.TryGetValue("ttlSeconds", out ttlSecondsText))
            {
                return null;
            }

            ManagedWorkbookCloseMarkerKind kind;
            DateTime createdUtc;
            int ttlSeconds;
            if (!Enum.TryParse(kindText, ignoreCase: true, result: out kind)
                || !DateTime.TryParse(createdUtcText, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out createdUtc)
                || !int.TryParse(ttlSecondsText, NumberStyles.Integer, CultureInfo.InvariantCulture, out ttlSeconds)
                || ttlSeconds <= 0
                || ttlSeconds > 30)
            {
                return null;
            }

            string workbookKeyBase64;
            values.TryGetValue("workbookKeyBase64", out workbookKeyBase64);

            return new ManagedWorkbookCloseMarker(kind, createdUtc.ToUniversalTime(), ttlSeconds, Decode(workbookKeyBase64));
        }

        private void TryDelete()
        {
            try
            {
                if (File.Exists(_markerPath))
                {
                    File.Delete(_markerPath);
                }
            }
            catch
            {
            }
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value));
            }
            catch
            {
                return string.Empty;
            }
        }

    }

    internal sealed class ManagedWorkbookCloseMarker
    {
        internal ManagedWorkbookCloseMarker(
            ManagedWorkbookCloseMarkerKind kind,
            DateTime createdUtc,
            int timeToLiveSeconds,
            string workbookKey)
        {
            Kind = kind;
            CreatedUtc = createdUtc;
            TimeToLiveSeconds = timeToLiveSeconds;
            WorkbookKey = workbookKey ?? string.Empty;
        }

        internal ManagedWorkbookCloseMarkerKind Kind { get; }

        internal DateTime CreatedUtc { get; }

        internal int TimeToLiveSeconds { get; }

        internal string WorkbookKey { get; }
    }

    internal enum ManagedWorkbookCloseMarkerReadStatus
    {
        NoMarker = 0,
        Valid = 1,
        Expired = 2,
        Invalid = 3
    }

    internal sealed class ManagedWorkbookCloseMarkerReadResult
    {
        private ManagedWorkbookCloseMarkerReadResult(
            ManagedWorkbookCloseMarkerReadStatus status,
            string markerPath,
            ManagedWorkbookCloseMarker marker,
            TimeSpan? age)
        {
            Status = status;
            MarkerPath = markerPath ?? string.Empty;
            Marker = marker;
            Age = age;
        }

        internal ManagedWorkbookCloseMarkerReadStatus Status { get; }

        internal string MarkerPath { get; }

        internal ManagedWorkbookCloseMarker Marker { get; }

        internal TimeSpan? Age { get; }

        internal bool IsValid
        {
            get { return Status == ManagedWorkbookCloseMarkerReadStatus.Valid && Marker != null; }
        }

        internal static ManagedWorkbookCloseMarkerReadResult NoMarker(string markerPath)
        {
            return new ManagedWorkbookCloseMarkerReadResult(ManagedWorkbookCloseMarkerReadStatus.NoMarker, markerPath, null, null);
        }

        internal static ManagedWorkbookCloseMarkerReadResult Valid(string markerPath, ManagedWorkbookCloseMarker marker, TimeSpan age)
        {
            return new ManagedWorkbookCloseMarkerReadResult(ManagedWorkbookCloseMarkerReadStatus.Valid, markerPath, marker, age);
        }

        internal static ManagedWorkbookCloseMarkerReadResult Expired(string markerPath, ManagedWorkbookCloseMarker marker, TimeSpan age)
        {
            return new ManagedWorkbookCloseMarkerReadResult(ManagedWorkbookCloseMarkerReadStatus.Expired, markerPath, marker, age);
        }

        internal static ManagedWorkbookCloseMarkerReadResult Invalid(string markerPath)
        {
            return new ManagedWorkbookCloseMarkerReadResult(ManagedWorkbookCloseMarkerReadStatus.Invalid, markerPath, null, null);
        }
    }
}
