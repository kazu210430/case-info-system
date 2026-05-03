using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelOpenWorkbookLocator
    {
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook> _getOpenKernelWorkbookOverride;
        private readonly Func<string, string> _resolveKernelWorkbookPathOverride;
        private readonly Func<string, Excel.Workbook> _findOpenWorkbookOverride;

        internal KernelOpenWorkbookLocator(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger,
            Func<Excel.Workbook> getOpenKernelWorkbookOverride = null,
            Func<string, string> resolveKernelWorkbookPathOverride = null,
            Func<string, Excel.Workbook> findOpenWorkbookOverride = null)
        {
            _application = application;
            _excelInteropService = excelInteropService;
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _getOpenKernelWorkbookOverride = getOpenKernelWorkbookOverride;
            _resolveKernelWorkbookPathOverride = resolveKernelWorkbookPathOverride;
            _findOpenWorkbookOverride = findOpenWorkbookOverride;
        }

        internal Excel.Workbook GetOpenKernelWorkbook()
        {
            if (_getOpenKernelWorkbookOverride != null)
            {
                return _getOpenKernelWorkbookOverride();
            }

            try
            {
                foreach (Excel.Workbook workbook in _application.Workbooks)
                {
                    if (IsKernelWorkbook(workbook))
                    {
                        return workbook;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("GetOpenKernelWorkbook failed.", ex);
            }

            return null;
        }

        internal Excel.Workbook ResolveKernelWorkbook(WorkbookContext context)
        {
            if (context == null)
            {
                return GetOpenKernelWorkbook();
            }

            Excel.Workbook contextWorkbook = context.Workbook;
            if (IsKernelWorkbook(contextWorkbook))
            {
                return contextWorkbook;
            }

            return ResolveKernelWorkbook(context.SystemRoot);
        }

        internal Excel.Workbook ResolveKernelWorkbook(string systemRoot)
        {
            string normalizedSystemRoot = _pathCompatibilityService.NormalizePath(systemRoot);
            string kernelPath = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: false,
                systemRoot: normalizedSystemRoot,
                resolvePath: root => ResolveKernelWorkbookPathCore(root));
            if (string.IsNullOrWhiteSpace(kernelPath))
            {
                return null;
            }

            return FindOpenWorkbookCore(kernelPath);
        }

        private string ResolveKernelWorkbookPathCore(string systemRoot)
        {
            return _resolveKernelWorkbookPathOverride != null
                ? _resolveKernelWorkbookPathOverride(systemRoot)
                : WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private Excel.Workbook FindOpenWorkbookCore(string workbookPath)
        {
            if (_findOpenWorkbookOverride != null)
            {
                return _findOpenWorkbookOverride(workbookPath);
            }

            return _excelInteropService == null ? null : _excelInteropService.FindOpenWorkbook(workbookPath);
        }

        private bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return WorkbookFileNameResolver.IsKernelWorkbookName(GetWorkbookNameCore(workbook));
        }

        private string GetWorkbookNameCore(Excel.Workbook workbook)
        {
            if (_excelInteropService != null)
            {
                return _excelInteropService.GetWorkbookName(workbook);
            }

            try
            {
                return workbook == null ? string.Empty : workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
