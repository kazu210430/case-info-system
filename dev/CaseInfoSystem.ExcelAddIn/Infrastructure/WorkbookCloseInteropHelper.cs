using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class WorkbookCloseInteropHelper
    {
        internal static void Close(Excel.Workbook workbook)
        {
            Close(workbook, null, null);
        }

        internal static void Close(Excel.Workbook workbook, Logger logger, string routeName)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            CloseCore(workbook, Type.Missing, "Type.Missing", logger, routeName);
        }

        internal static void CloseWithoutSave(Excel.Workbook workbook)
        {
            CloseWithoutSave(workbook, null, null);
        }

        internal static void CloseWithoutSave(Excel.Workbook workbook, Logger logger, string routeName)
        {
            CloseWithoutSaveCore(workbook, logger, routeName);
        }

        internal static void CloseReadOnlyWithoutSave(Excel.Workbook workbook, Logger logger, string routeName)
        {
            CloseWithoutSaveCore(workbook, logger, NormalizeRouteName(routeName, "read-only-without-save"));
        }

        internal static void CloseOwnedWorkbookWithoutSave(Excel.Workbook workbook, Logger logger, string routeName)
        {
            CloseWithoutSaveCore(workbook, logger, NormalizeRouteName(routeName, "owned-workbook-without-save"));
        }

        private static void CloseWithoutSaveCore(Excel.Workbook workbook, Logger logger, string routeName)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            CloseCore(workbook, false, "false", logger, routeName);
        }

        private static void CloseCore(Excel.Workbook workbook, object saveChanges, string saveChangesText, Logger logger, string routeName)
        {
            string safeRouteName = routeName ?? string.Empty;
            string workbookName = SafeWorkbookName(workbook);
            string workbookPath = SafeWorkbookPath(workbook);
            string safeSaveChangesText = saveChangesText ?? string.Empty;

            Debug(
                logger,
                "Workbook close starting. route="
                + safeRouteName
                + ", workbookName="
                + workbookName
                + ", workbookPath="
                + workbookPath
                + ", saveChanges="
                + safeSaveChangesText
                + ", closeHelperContract=caller-does-not-read-closed-workbook");

            try
            {
                workbook.Close(saveChanges, Type.Missing, Type.Missing);
            }
            catch (Exception exception)
            {
                Error(
                    logger,
                    "Workbook close failed. route="
                    + safeRouteName
                    + ", workbookName="
                    + workbookName
                    + ", workbookPath="
                    + workbookPath
                    + ", saveChanges="
                    + safeSaveChangesText,
                    exception);
                throw;
            }

            Debug(
                logger,
                "Workbook close completed. route="
                + safeRouteName
                + ", workbookName="
                + workbookName
                + ", workbookPath="
                + workbookPath
                + ", saveChanges="
                + safeSaveChangesText
                + ", closeSucceeded=true"
                + ", closeHelperContract=caller-does-not-read-closed-workbook");
        }

        private static string NormalizeRouteName(string routeName, string defaultRouteName)
        {
            return string.IsNullOrWhiteSpace(routeName)
                ? defaultRouteName ?? string.Empty
                : routeName;
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

        private static string SafeWorkbookPath(Excel.Workbook workbook)
        {
            try
            {
                return workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void Debug(Logger logger, string message)
        {
            if (logger == null)
            {
                return;
            }

            try
            {
                logger.Debug(nameof(WorkbookCloseInteropHelper), message);
            }
            catch
            {
            }
        }

        private static void Error(Logger logger, string message, Exception exception)
        {
            if (logger == null)
            {
                return;
            }

            try
            {
                logger.Error(message, exception);
            }
            catch
            {
            }
        }
    }
}
