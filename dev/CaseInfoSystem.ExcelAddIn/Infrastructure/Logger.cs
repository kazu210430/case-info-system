using System;
using System.Globalization;
using System.Threading;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class Logger
    {
        private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
        private readonly Action<string> _writeTrace;

        /// <summary>
        internal Logger(Action<string> writeTrace)
        {
            _writeTrace = writeTrace ?? throw new ArgumentNullException(nameof(writeTrace));
        }

        /// <summary>
        internal void Info(string message)
        {
            _writeTrace(WithKernelFlickerTraceId(message));
        }

        /// <summary>
        internal void Debug(string procedureName, string message)
        {
            string safeProcedureName = procedureName ?? string.Empty;
            string safeMessage = message ?? string.Empty;
            _writeTrace(WithKernelFlickerTraceId("DEBUG: " + safeProcedureName + " " + safeMessage));
        }

        /// <summary>
        internal void Warn(string message)
        {
            _writeTrace(WithKernelFlickerTraceId("WARN: " + (message ?? string.Empty)));
        }

        /// <summary>
        internal void Error(string context, Exception exception)
        {
            string safeContext = context ?? string.Empty;
            if (exception == null)
            {
                _writeTrace(WithKernelFlickerTraceId("ERROR: " + safeContext));
                return;
            }

            _writeTrace(WithKernelFlickerTraceId("ERROR: " + safeContext + " " + exception.GetType().Name + ": " + exception.Message));
        }

        /// <summary>
        internal void Error(string procedureName, long errorNumber, string errorDescription)
        {
            string safeProcedureName = procedureName ?? string.Empty;
            string safeDescription = errorDescription ?? string.Empty;
            _writeTrace(WithKernelFlickerTraceId("ERROR: " + safeProcedureName + " Err=" + errorNumber.ToString() + " " + safeDescription));
        }

        private static string WithKernelFlickerTraceId(string message)
        {
            string safeMessage = message ?? string.Empty;
            string traceId = KernelFlickerTraceContext.CurrentTraceId;
            if (string.IsNullOrWhiteSpace(traceId))
            {
                return safeMessage;
            }

            string targetPrefix = KernelFlickerTracePrefix + " ";
            if (safeMessage.StartsWith(targetPrefix, StringComparison.Ordinal))
            {
                return KernelFlickerTracePrefix + " traceId=" + traceId + " " + safeMessage.Substring(targetPrefix.Length);
            }

            return safeMessage;
        }
    }

    internal static class KernelFlickerTraceContext
    {
        private static readonly AsyncLocal<string> CurrentTraceIdSlot = new AsyncLocal<string>();
        private static int _traceSequence;

        internal static string CurrentTraceId => CurrentTraceIdSlot.Value ?? string.Empty;

        internal static string BeginNewTrace()
        {
            string traceId = "KF-"
                + Interlocked.Increment(ref _traceSequence).ToString("D6", CultureInfo.InvariantCulture);
            CurrentTraceIdSlot.Value = traceId;
            return traceId;
        }

        internal static void SetCurrentTrace(string traceId)
        {
            CurrentTraceIdSlot.Value = traceId ?? string.Empty;
        }
    }
}
