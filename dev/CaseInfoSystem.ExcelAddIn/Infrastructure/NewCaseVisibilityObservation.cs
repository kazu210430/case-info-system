using System;
using System.Collections.Generic;
using System.Globalization;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class NewCaseVisibilityObservation
    {
        private const string LogPrefix = "[NewCaseVisibilityObservation]";
        private static readonly object SyncRoot = new object();
        private static readonly Dictionary<string, Session> Sessions = new Dictionary<string, Session>(StringComparer.OrdinalIgnoreCase);
        private static readonly TimeSpan SessionTtl = TimeSpan.FromMinutes(2);
        private static int _sessionSequence;

        private sealed class Session
        {
            internal Session(KernelCaseCreationMode mode, string sessionId, string primaryCaseWorkbookPath)
            {
                Mode = mode;
                SessionId = sessionId ?? string.Empty;
                PrimaryCaseWorkbookPath = primaryCaseWorkbookPath ?? string.Empty;
                CreatedUtc = DateTime.UtcNow;
                Keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            internal KernelCaseCreationMode Mode { get; }

            internal string SessionId { get; }

            internal string PrimaryCaseWorkbookPath { get; }

            internal DateTime CreatedUtc { get; }

            internal HashSet<string> Keys { get; }
        }

        internal static void Begin(KernelCaseCreationMode mode, string caseWorkbookPath)
        {
            string normalizedPath = NormalizeKey(caseWorkbookPath);
            if (string.IsNullOrWhiteSpace(normalizedPath))
            {
                return;
            }

            lock (SyncRoot)
            {
                if (TryGetLiveSessionLocked(normalizedPath, out Session existingSession))
                {
                    return;
                }

                string sessionId = "NCO-" + (++_sessionSequence).ToString("D4", CultureInfo.InvariantCulture);
                Session session = new Session(mode, sessionId, normalizedPath);
                session.Keys.Add(normalizedPath);
                Sessions[normalizedPath] = session;
            }
        }

        internal static void AttachAlias(string existingCaseWorkbookPath, string aliasWorkbookPath)
        {
            string normalizedExistingPath = NormalizeKey(existingCaseWorkbookPath);
            string normalizedAliasPath = NormalizeKey(aliasWorkbookPath);
            if (string.IsNullOrWhiteSpace(normalizedExistingPath) || string.IsNullOrWhiteSpace(normalizedAliasPath))
            {
                return;
            }

            lock (SyncRoot)
            {
                if (!TryGetLiveSessionLocked(normalizedExistingPath, out Session session)
                    && !TryGetLiveSessionLocked(normalizedAliasPath, out session))
                {
                    return;
                }

                session.Keys.Add(normalizedAliasPath);
                Sessions[normalizedAliasPath] = session;
            }
        }

        internal static void Complete(string caseWorkbookPath)
        {
            string normalizedPath = NormalizeKey(caseWorkbookPath);
            if (string.IsNullOrWhiteSpace(normalizedPath))
            {
                return;
            }

            lock (SyncRoot)
            {
                if (TryGetLiveSessionLocked(normalizedPath, out Session session))
                {
                    RemoveSessionLocked(session);
                }
            }
        }

        internal static void Log(
            Logger logger,
            ExcelInteropService excelInteropService,
            Excel.Application application,
            Excel.Workbook workbook,
            Excel.Window window,
            string stepName,
            string source,
            string caseWorkbookPath = null,
            string detail = null)
        {
            if (logger == null)
            {
                return;
            }

            try
            {
                Session session = ResolveTrackedSession(caseWorkbookPath, excelInteropService, workbook);
                if (session == null)
                {
                    return;
                }

                Excel.Application resolvedApplication = ResolveApplication(application, workbook);
                Excel.Workbook activeWorkbook = ResolveActiveWorkbook(resolvedApplication, excelInteropService);
                Excel.Window activeWindow = ResolveActiveWindow(resolvedApplication, excelInteropService);

                logger.Info(
                    LogPrefix
                    + " timestamp=" + DateTimeOffset.Now.ToString("O", CultureInfo.InvariantCulture)
                    + " sessionId=" + session.SessionId
                    + ", mode=" + session.Mode.ToString()
                    + ", step=" + (stepName ?? string.Empty)
                    + ", source=" + (source ?? string.Empty)
                    + ", caseWorkbookPath=\"" + session.PrimaryCaseWorkbookPath + "\""
                    + ", appHwnd=\"" + SafeApplicationHwnd(resolvedApplication) + "\""
                    + ", appVisible=" + SafeApplicationVisible(resolvedApplication)
                    + ", screenUpdating=" + SafeScreenUpdating(resolvedApplication)
                    + ", displayAlerts=" + SafeDisplayAlerts(resolvedApplication)
                    + ", enableEvents=" + SafeEnableEvents(resolvedApplication)
                    + ", workbook=" + FormatWorkbookDescriptor(excelInteropService, workbook)
                    + ", workbookWindow=" + FormatWorkbookWindowDescriptor(workbook, window)
                    + ", eventWindow=" + FormatWindowDescriptor(window)
                    + ", activeWorkbook=" + FormatWorkbookDescriptor(excelInteropService, activeWorkbook)
                    + ", activeWindow=" + FormatWindowDescriptor(activeWindow)
                    + FormatDetail(detail));
            }
            catch
            {
            }
        }

        private static Session ResolveTrackedSession(string explicitCaseWorkbookPath, ExcelInteropService excelInteropService, Excel.Workbook workbook)
        {
            string normalizedPath = NormalizeKey(explicitCaseWorkbookPath);
            if (!string.IsNullOrWhiteSpace(normalizedPath))
            {
                Session explicitSession = TryGetLiveSession(normalizedPath);
                if (explicitSession != null)
                {
                    return explicitSession;
                }
            }

            string workbookFullName = NormalizeKey(SafeWorkbookFullName(excelInteropService, workbook));
            if (string.IsNullOrWhiteSpace(workbookFullName))
            {
                return null;
            }

            return TryGetLiveSession(workbookFullName);
        }

        private static Session TryGetLiveSession(string normalizedKey)
        {
            lock (SyncRoot)
            {
                return TryGetLiveSessionLocked(normalizedKey, out Session session) ? session : null;
            }
        }

        private static bool TryGetLiveSessionLocked(string normalizedKey, out Session session)
        {
            session = null;
            if (string.IsNullOrWhiteSpace(normalizedKey) || !Sessions.TryGetValue(normalizedKey, out Session candidate))
            {
                return false;
            }

            if (IsExpired(candidate))
            {
                RemoveSessionLocked(candidate);
                return false;
            }

            session = candidate;
            return true;
        }

        private static bool IsExpired(Session session)
        {
            return session == null || DateTime.UtcNow - session.CreatedUtc > SessionTtl;
        }

        private static void RemoveSessionLocked(Session session)
        {
            if (session == null)
            {
                return;
            }

            foreach (string key in session.Keys)
            {
                Sessions.Remove(key);
            }
        }

        private static string NormalizeKey(string key)
        {
            return string.IsNullOrWhiteSpace(key) ? string.Empty : key.Trim();
        }

        private static Excel.Application ResolveApplication(Excel.Application application, Excel.Workbook workbook)
        {
            if (application != null)
            {
                return application;
            }

            try
            {
                return workbook == null ? null : workbook.Application;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Workbook ResolveActiveWorkbook(Excel.Application application, ExcelInteropService excelInteropService)
        {
            try
            {
                if (excelInteropService != null)
                {
                    return excelInteropService.GetActiveWorkbook();
                }
            }
            catch
            {
            }

            try
            {
                return application == null ? null : application.ActiveWorkbook;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Window ResolveActiveWindow(Excel.Application application, ExcelInteropService excelInteropService)
        {
            try
            {
                if (excelInteropService != null)
                {
                    return excelInteropService.GetActiveWindow();
                }
            }
            catch
            {
            }

            try
            {
                return application == null ? null : application.ActiveWindow;
            }
            catch
            {
                return null;
            }
        }

        private static string FormatWorkbookDescriptor(ExcelInteropService excelInteropService, Excel.Workbook workbook)
        {
            return "full=\""
                + SafeWorkbookFullName(excelInteropService, workbook)
                + "\",name=\""
                + SafeWorkbookName(workbook)
                + "\"";
        }

        private static string FormatWorkbookWindowDescriptor(Excel.Workbook workbook, Excel.Window explicitWindow)
        {
            if (explicitWindow != null)
            {
                return FormatWindowDescriptor(explicitWindow);
            }

            Excel.Windows workbookWindows = null;
            Excel.Window workbookWindow = null;
            try
            {
                workbookWindows = workbook == null ? null : workbook.Windows;
                if (workbookWindows == null || workbookWindows.Count < 1)
                {
                    return FormatWindowDescriptor(null);
                }

                workbookWindow = workbookWindows[1];
                return FormatWindowDescriptor(workbookWindow);
            }
            catch
            {
                return FormatWindowDescriptor(null);
            }
            finally
            {
                ComObjectReleaseService.Release(workbookWindow);
                ComObjectReleaseService.Release(workbookWindows);
            }
        }

        private static string FormatWindowDescriptor(Excel.Window window)
        {
            return "hwnd=\""
                + SafeWindowHwnd(window)
                + "\",caption=\""
                + SafeWindowCaption(window)
                + "\",visible=\""
                + SafeWindowVisible(window)
                + "\",state=\""
                + SafeWindowState(window)
                + "\"";
        }

        private static string FormatDetail(string detail)
        {
            return string.IsNullOrWhiteSpace(detail)
                ? string.Empty
                : ", detail=\"" + detail.Trim() + "\"";
        }

        private static string SafeWorkbookFullName(ExcelInteropService excelInteropService, Excel.Workbook workbook)
        {
            try
            {
                if (excelInteropService != null)
                {
                    return excelInteropService.GetWorkbookFullName(workbook) ?? string.Empty;
                }
            }
            catch
            {
            }

            try
            {
                return workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWorkbookName(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null ? string.Empty : workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeApplicationHwnd(Excel.Application application)
        {
            try
            {
                return application == null
                    ? string.Empty
                    : Convert.ToString(application.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeApplicationVisible(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : application.Visible.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeScreenUpdating(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : application.ScreenUpdating.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeDisplayAlerts(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : application.DisplayAlerts.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeEnableEvents(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : application.EnableEvents.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowHwnd(Excel.Window window)
        {
            try
            {
                return window == null
                    ? string.Empty
                    : Convert.ToString(window.Hwnd, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowCaption(Excel.Window window)
        {
            try
            {
                if (window == null)
                {
                    return string.Empty;
                }

                dynamic lateBoundWindow = window;
                return Convert.ToString(lateBoundWindow.Caption, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowVisible(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.Visible.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeWindowState(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : window.WindowState.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
