using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelWorkbookResolver
    {
        private const string SystemRootPropName = "SYSTEM_ROOT";

        internal static string ResolveSystemRootFromAvailableWorkbooks(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger,
            Func<Excel.Workbook, bool> isKernelWorkbook)
        {
            string systemRoot = GetSystemRootFromWorkbook(application.ActiveWorkbook, excelInteropService, pathCompatibilityService, logger, isKernelWorkbook);
            if (!string.IsNullOrWhiteSpace(systemRoot))
            {
                return systemRoot;
            }

            foreach (Excel.Workbook workbook in application.Workbooks)
            {
                systemRoot = GetSystemRootFromWorkbook(workbook, excelInteropService, pathCompatibilityService, logger, isKernelWorkbook);
                if (!string.IsNullOrWhiteSpace(systemRoot))
                {
                    return systemRoot;
                }
            }

            return string.Empty;
        }

        internal static string GetSystemRootFromWorkbook(
            Excel.Workbook workbook,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger,
            Func<Excel.Workbook, bool> isKernelWorkbook)
        {
            string systemRoot = NormalizePath(excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropName));
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                string workbookPathFallback = TryGetKernelWorkbookDirectory(workbook, excelInteropService, pathCompatibilityService, logger, isKernelWorkbook);
                if (string.IsNullOrWhiteSpace(workbookPathFallback))
                {
                    return string.Empty;
                }

                logger.Info("GetSystemRootFromWorkbook fallback used workbook directory. workbook=" + excelInteropService.GetWorkbookFullName(workbook));
                return workbookPathFallback;
            }

            string kernelWorkbookPath = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, pathCompatibilityService);
            return string.IsNullOrWhiteSpace(kernelWorkbookPath) ? string.Empty : systemRoot;
        }

        internal static string TryGetKernelWorkbookDirectory(
            Excel.Workbook workbook,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger,
            Func<Excel.Workbook, bool> isKernelWorkbook)
        {
            if (!isKernelWorkbook(workbook))
            {
                return string.Empty;
            }

            string workbookFullName = NormalizePath(excelInteropService.GetWorkbookFullName(workbook));
            if (string.IsNullOrWhiteSpace(workbookFullName) || !File.Exists(workbookFullName))
            {
                return string.Empty;
            }

            string workbookDirectory;
            try
            {
                workbookDirectory = Path.GetDirectoryName(workbookFullName);
            }
            catch (Exception ex)
            {
                logger.Error("TryGetKernelWorkbookDirectory Path.GetDirectoryName failed.", ex);
                return string.Empty;
            }

            workbookDirectory = NormalizePath(workbookDirectory);
            if (string.IsNullOrWhiteSpace(workbookDirectory))
            {
                return string.Empty;
            }

            string kernelWorkbookPath = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(workbookDirectory, pathCompatibilityService);
            return string.IsNullOrWhiteSpace(kernelWorkbookPath) ? string.Empty : workbookDirectory;
        }

        internal static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar);
            }
            catch
            {
                return path.Trim();
            }
        }
    }
}
