using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class Logger
    {
        private readonly Action<string> _writeTrace;

        internal Logger(Action<string> writeTrace)
        {
            _writeTrace = writeTrace ?? (_ => { });
        }

        internal void Info(string message)
        {
            _writeTrace(message ?? string.Empty);
        }

        internal void Warn(string message)
        {
            _writeTrace(message ?? string.Empty);
        }

        internal void Error(string context, Exception exception)
        {
            _writeTrace((context ?? string.Empty) + " " + (exception?.Message ?? string.Empty));
        }
    }

    internal sealed class PathCompatibilityService
    {
        internal string NormalizePath(string path)
        {
            return (path ?? string.Empty).Trim().Replace("/", "\\");
        }

        internal bool FileExistsSafe(string path)
        {
            return File.Exists(NormalizePath(path));
        }

        internal string GetFileNameFromPath(string fullPath)
        {
            return Path.GetFileName(NormalizePath(fullPath)) ?? string.Empty;
        }

        internal string GetParentFolderPath(string fullPath)
        {
            return Path.GetDirectoryName(NormalizePath(fullPath)) ?? string.Empty;
        }

        internal string CombinePath(string left, string right)
        {
            if (string.IsNullOrWhiteSpace(left))
            {
                return NormalizePath(right);
            }

            if (string.IsNullOrWhiteSpace(right))
            {
                return NormalizePath(left);
            }

            return NormalizePath(Path.Combine(NormalizePath(left), NormalizePath(right)));
        }
    }

    internal static class WorkbookFileNameResolver
    {
        internal const string KernelWorkbookBaseName = "案件情報System_Kernel";
        internal const string BaseWorkbookBaseName = "案件情報System_Base";

        internal static string ResolveExistingKernelWorkbookPath(string systemRoot, PathCompatibilityService pathCompatibilityService)
        {
            string root = pathCompatibilityService?.NormalizePath(systemRoot) ?? string.Empty;
            if (root.Length == 0)
            {
                return string.Empty;
            }

            return pathCompatibilityService.CombinePath(root, KernelWorkbookBaseName + ".xlsx");
        }

        internal static string BuildBaseWorkbookName(string extension)
        {
            string normalizedExtension = string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                ? ".xlsx"
                : ".xlsm";
            return BaseWorkbookBaseName + normalizedExtension;
        }
    }

    internal sealed class ExcelInteropService
    {
        private readonly Excel.Application _application;
        private readonly PathCompatibilityService _pathCompatibilityService;

        internal ExcelInteropService(Excel.Application application, Logger logger, PathCompatibilityService pathCompatibilityService)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
        }

        internal string TryGetDocumentProperty(Excel.Workbook workbook, string propertyName)
        {
            if (workbook?.CustomDocumentProperties is DocumentProperties properties)
            {
                return Convert.ToString(properties[propertyName]?.Value, CultureInfo.InvariantCulture) ?? string.Empty;
            }

            return string.Empty;
        }

        internal void SetDocumentProperty(Excel.Workbook workbook, string propertyName, string value)
        {
            if (workbook.CustomDocumentProperties is not DocumentProperties properties)
            {
                properties = new DocumentProperties();
                workbook.CustomDocumentProperties = properties;
            }

            DocumentProperty property = properties[propertyName];
            if (property == null)
            {
                properties.Add(propertyName, false, 4, value ?? string.Empty);
            }
            else
            {
                property.Value = value ?? string.Empty;
            }
        }

        internal string GetWorkbookName(Excel.Workbook workbook)
        {
            return workbook?.Name ?? string.Empty;
        }

        internal string GetWorkbookFullName(Excel.Workbook workbook)
        {
            return workbook?.FullName ?? string.Empty;
        }

        internal string GetWorkbookPath(Excel.Workbook workbook)
        {
            return workbook?.Path ?? string.Empty;
        }

        internal Excel.Workbook FindOpenWorkbook(string workbookFullName)
        {
            string normalizedTarget = _pathCompatibilityService.NormalizePath(workbookFullName);
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                string fullName = _pathCompatibilityService.NormalizePath(workbook?.FullName);
                if (string.Equals(fullName, normalizedTarget, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }

            return null;
        }

        internal Excel.Worksheet FindWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(sheetCodeName))
            {
                return null;
            }

            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet?.CodeName, sheetCodeName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            return null;
        }
    }
}
