using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal delegate void CancelableEventBoundaryAction(ref bool cancel);

    internal static class EventBoundaryGuard
    {
        internal static void Execute(Logger logger, string eventName, Action action)
        {
            try
            {
                action?.Invoke();
            }
            catch (Exception ex)
            {
                logger?.Error((eventName ?? string.Empty) + " failed.", ex);
            }
        }

        internal static void ExecuteCancelable(Logger logger, string eventName, ref bool cancel, Action action)
        {
            try
            {
                action?.Invoke();
            }
            catch (Exception ex)
            {
                cancel = true;
                logger?.Error((eventName ?? string.Empty) + " failed.", ex);
            }
        }

        internal static void ExecuteCancelable(Logger logger, string eventName, ref bool cancel, CancelableEventBoundaryAction action)
        {
            try
            {
                action?.Invoke(ref cancel);
            }
            catch (Exception ex)
            {
                cancel = true;
                logger?.Error((eventName ?? string.Empty) + " failed.", ex);
            }
        }
    }
}
