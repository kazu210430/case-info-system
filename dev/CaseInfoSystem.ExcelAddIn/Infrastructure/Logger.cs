using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class Logger
    {
        private readonly Action<string> _writeTrace;

        /// <summary>
        internal Logger(Action<string> writeTrace)
        {
            _writeTrace = writeTrace ?? throw new ArgumentNullException(nameof(writeTrace));
        }

        /// <summary>
        internal void Info(string message)
        {
            _writeTrace(message ?? string.Empty);
        }

        /// <summary>
        internal void Debug(string procedureName, string message)
        {
            string safeProcedureName = procedureName ?? string.Empty;
            string safeMessage = message ?? string.Empty;
            _writeTrace("DEBUG: " + safeProcedureName + " " + safeMessage);
        }

        /// <summary>
        internal void Warn(string message)
        {
            _writeTrace("WARN: " + (message ?? string.Empty));
        }

        /// <summary>
        internal void Error(string context, Exception exception)
        {
            string safeContext = context ?? string.Empty;
            if (exception == null)
            {
                _writeTrace("ERROR: " + safeContext);
                return;
            }

            _writeTrace("ERROR: " + safeContext + " " + exception.GetType().Name + ": " + exception.Message);
        }

        /// <summary>
        internal void Error(string procedureName, long errorNumber, string errorDescription)
        {
            string safeProcedureName = procedureName ?? string.Empty;
            string safeDescription = errorDescription ?? string.Empty;
            _writeTrace("ERROR: " + safeProcedureName + " Err=" + errorNumber.ToString() + " " + safeDescription);
        }
    }
}
