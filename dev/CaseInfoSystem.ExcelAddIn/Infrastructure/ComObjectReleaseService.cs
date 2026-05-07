using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class ComObjectReleaseService
    {
        internal sealed class ComObjectReleaseServiceTestHooks
        {
            internal Func<object, bool> IsComObject { get; set; }

            internal Func<object, int> ReleaseComObject { get; set; }

            internal Func<object, int> FinalReleaseComObject { get; set; }

            internal Action<string> DebugWriteLine { get; set; }
        }

        internal static void Release(object comObject)
        {
            ReleaseCore(comObject, finalRelease: false, logger: null, context: null, callerMemberName: null, testHooks: null);
        }

        internal static void Release(object comObject, Logger logger, string context = null, [CallerMemberName] string callerMemberName = null)
        {
            ReleaseCore(comObject, finalRelease: false, logger, context, callerMemberName, testHooks: null);
        }

        internal static void Release(object comObject, Logger logger, string context, ComObjectReleaseServiceTestHooks testHooks)
        {
            ReleaseCore(comObject, finalRelease: false, logger, context, callerMemberName: null, testHooks);
        }

        internal static void FinalRelease(object comObject)
        {
            ReleaseCore(comObject, finalRelease: true, logger: null, context: null, callerMemberName: null, testHooks: null);
        }

        internal static void FinalRelease(object comObject, Logger logger, string context = null, [CallerMemberName] string callerMemberName = null)
        {
            ReleaseCore(comObject, finalRelease: true, logger, context, callerMemberName, testHooks: null);
        }

        internal static void FinalRelease(object comObject, Logger logger, string context, ComObjectReleaseServiceTestHooks testHooks)
        {
            ReleaseCore(comObject, finalRelease: true, logger, context, callerMemberName: null, testHooks);
        }

        private static void ReleaseCore(
            object comObject,
            bool finalRelease,
            Logger logger,
            string context,
            string callerMemberName,
            ComObjectReleaseServiceTestHooks testHooks)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (!IsComObject(comObject, testHooks))
                {
                    return;
                }

                if (finalRelease)
                {
                    FinalReleaseComObject(comObject, testHooks);
                    return;
                }

                ReleaseComObject(comObject, testHooks);
            }
            catch (Exception ex)
            {
                string operation = finalRelease ? "FinalRelease" : "Release";
                WriteDebugLine("ComObjectReleaseService." + operation + " failed: " + ex, testHooks);
                TryWarn(logger, operation, comObject, ResolveContext(context, callerMemberName), ex);
            }
        }

        private static bool IsComObject(object comObject, ComObjectReleaseServiceTestHooks testHooks)
        {
            if (testHooks != null && testHooks.IsComObject != null)
            {
                return testHooks.IsComObject(comObject);
            }

            return Marshal.IsComObject(comObject);
        }

        private static void ReleaseComObject(object comObject, ComObjectReleaseServiceTestHooks testHooks)
        {
            if (testHooks != null && testHooks.ReleaseComObject != null)
            {
                testHooks.ReleaseComObject(comObject);
                return;
            }

            Marshal.ReleaseComObject(comObject);
        }

        private static void FinalReleaseComObject(object comObject, ComObjectReleaseServiceTestHooks testHooks)
        {
            if (testHooks != null && testHooks.FinalReleaseComObject != null)
            {
                testHooks.FinalReleaseComObject(comObject);
                return;
            }

            Marshal.FinalReleaseComObject(comObject);
        }

        private static void WriteDebugLine(string message, ComObjectReleaseServiceTestHooks testHooks)
        {
            if (testHooks != null && testHooks.DebugWriteLine != null)
            {
                testHooks.DebugWriteLine(message);
                return;
            }

            Debug.WriteLine(message);
        }

        private static void TryWarn(Logger logger, string operation, object comObject, string context, Exception exception)
        {
            if (logger == null)
            {
                return;
            }

            try
            {
                logger.Warn(
                    "COM cleanup release failed. operation="
                    + (operation ?? string.Empty)
                    + ", targetType="
                    + GetTargetTypeName(comObject)
                    + ", context="
                    + SanitizeForSingleLine(context)
                    + ", exceptionType="
                    + (exception == null ? string.Empty : exception.GetType().Name)
                    + ", message="
                    + SanitizeForSingleLine(exception == null ? string.Empty : exception.Message));
            }
            catch
            {
                // Logging failures must not break best-effort cleanup.
            }
        }

        private static string ResolveContext(string context, string callerMemberName)
        {
            if (!string.IsNullOrWhiteSpace(context))
            {
                return context;
            }

            return string.IsNullOrWhiteSpace(callerMemberName) ? "(unspecified)" : callerMemberName;
        }

        private static string GetTargetTypeName(object comObject)
        {
            try
            {
                return comObject == null
                    ? string.Empty
                    : (comObject.GetType().FullName ?? comObject.GetType().Name ?? string.Empty);
            }
            catch
            {
                return "(type-unavailable)";
            }
        }

        private static string SanitizeForSingleLine(string value)
        {
            string safeValue = value ?? string.Empty;
            return safeValue
                .Replace("\r\n", " | ")
                .Replace("\n", " | ")
                .Replace("\r", " | ");
        }
    }
}
