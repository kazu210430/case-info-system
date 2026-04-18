using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class TransientPaneSuppressionService
    {
        private readonly IExcelInteropService _excelInteropService;
        private readonly IPathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;
        private readonly HashSet<string> _suppressedWorkbookKeys;

        internal TransientPaneSuppressionService(
            IExcelInteropService excelInteropService,
            IPathCompatibilityService pathCompatibilityService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _suppressedWorkbookKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        internal void SuppressPath(string workbookPath, string reason)
        {
            string normalized = Normalize(workbookPath);
            if (normalized.Length == 0)
            {
                return;
            }

            if (_suppressedWorkbookKeys.Add(normalized))
            {
                _logger.Info("TaskPane suppression registered. path=" + normalized + ", reason=" + (reason ?? string.Empty));
            }
        }

        internal void ReleasePath(string workbookPath, string reason)
        {
            string normalized = Normalize(workbookPath);
            if (normalized.Length == 0)
            {
                return;
            }

            if (_suppressedWorkbookKeys.Remove(normalized))
            {
                _logger.Info("TaskPane suppression released. path=" + normalized + ", reason=" + (reason ?? string.Empty));
            }
        }

        internal void ReleaseWorkbook(Excel.Workbook workbook, string reason)
        {
            if (workbook == null)
            {
                return;
            }

            ReleasePath(_excelInteropService.GetWorkbookFullName(workbook), reason);
        }

        internal bool IsSuppressed(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            return IsSuppressedPath(_excelInteropService.GetWorkbookFullName(workbook));
        }

        internal bool IsSuppressedPath(string workbookPath)
        {
            string normalized = Normalize(workbookPath);
            return normalized.Length > 0 && _suppressedWorkbookKeys.Contains(normalized);
        }

        private string Normalize(string workbookPath)
        {
            return _pathCompatibilityService.NormalizePath(workbookPath);
        }
    }
}
