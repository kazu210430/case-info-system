using System;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookStateService
    {
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook> _getOpenKernelWorkbookOverride;
        private readonly Func<string, string> _resolveKernelWorkbookPathOverride;
        private readonly Func<string, Excel.Workbook> _findOpenWorkbookOverride;
        private readonly Func<Excel.Workbook, bool> _hasOtherVisibleWorkbookOverride;
        private readonly Func<Excel.Workbook, bool> _hasOtherWorkbookOverride;

        internal KernelWorkbookStateService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger,
            Func<Excel.Workbook> getOpenKernelWorkbookOverride = null,
            Func<string, string> resolveKernelWorkbookPathOverride = null,
            Func<string, Excel.Workbook> findOpenWorkbookOverride = null,
            Func<Excel.Workbook, bool> hasOtherVisibleWorkbookOverride = null,
            Func<Excel.Workbook, bool> hasOtherWorkbookOverride = null)
        {
            _application = application;
            _excelInteropService = excelInteropService;
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _getOpenKernelWorkbookOverride = getOpenKernelWorkbookOverride;
            _resolveKernelWorkbookPathOverride = resolveKernelWorkbookPathOverride;
            _findOpenWorkbookOverride = findOpenWorkbookOverride;
            _hasOtherVisibleWorkbookOverride = hasOtherVisibleWorkbookOverride;
            _hasOtherWorkbookOverride = hasOtherWorkbookOverride;
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

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return WorkbookFileNameResolver.IsKernelWorkbookName(GetWorkbookNameCore(workbook));
        }

        internal Excel.Workbook ResolveKernelWorkbook(WorkbookContext context)
        {
            Excel.Workbook openKernelWorkbook = GetOpenKernelWorkbook();
            string kernelPath = KernelWorkbookResolutionPolicy.ResolveKernelWorkbookPath(
                hasOpenKernelWorkbook: openKernelWorkbook != null,
                systemRoot: context == null ? string.Empty : context.SystemRoot,
                resolvePath: root => ResolveKernelWorkbookPathCore(root));

            if (openKernelWorkbook != null)
            {
                return openKernelWorkbook;
            }

            if (string.IsNullOrWhiteSpace(kernelPath))
            {
                return null;
            }

            return FindOpenWorkbookCore(kernelPath);
        }

        internal bool ShouldShowHomeOnStartup(Excel.Workbook startupWorkbook = null)
        {
            bool hasExplicitKernelStartupContext = HasExplicitKernelStartupContext(startupWorkbook);
            bool hasKernelWorkbookContext = hasExplicitKernelStartupContext && HasKernelWorkbookContext();
            bool isStartupWorkbookKernel = hasExplicitKernelStartupContext && hasKernelWorkbookContext && IsKernelWorkbook(startupWorkbook);
            bool hasVisibleNonKernelWorkbook = hasExplicitKernelStartupContext && hasKernelWorkbookContext && !isStartupWorkbookKernel && HasVisibleNonKernelWorkbook();
            return KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext,
                hasKernelWorkbookContext,
                isStartupWorkbookKernel,
                hasVisibleNonKernelWorkbook);
        }

        internal string DescribeStartupState()
        {
            string activeWorkbookName = "(null)";
            bool activeIsKernel = false;
            bool hasOpenKernelWorkbook = false;
            bool hasVisibleNonKernelWorkbook = false;

            try
            {
                Excel.Workbook activeWorkbook = _application.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    activeWorkbookName = GetWorkbookNameCore(activeWorkbook);
                    activeIsKernel = IsKernelWorkbook(activeWorkbook);
                }
            }
            catch
            {
                activeWorkbookName = "(error)";
            }

            try
            {
                hasOpenKernelWorkbook = GetOpenKernelWorkbook() != null;
                hasVisibleNonKernelWorkbook = HasVisibleNonKernelWorkbook();
            }
            catch
            {
            }

            return "activeWorkbook="
                + activeWorkbookName
                + ", activeIsKernel="
                + activeIsKernel
                + ", hasOpenKernelWorkbook="
                + hasOpenKernelWorkbook
                + ", hasVisibleNonKernelWorkbook="
                + hasVisibleNonKernelWorkbook;
        }

        internal bool HasOtherVisibleWorkbook(Excel.Workbook workbookToIgnore)
        {
            if (_hasOtherVisibleWorkbookOverride != null)
            {
                return _hasOtherVisibleWorkbookOverride(workbookToIgnore);
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                {
                    continue;
                }

                if (workbook.Windows.Count > 0 && workbook.Windows.Cast<Excel.Window>().Any(window => window.Visible))
                {
                    return true;
                }
            }

            return false;
        }

        internal bool HasOtherWorkbook(Excel.Workbook workbookToIgnore)
        {
            if (_hasOtherWorkbookOverride != null)
            {
                return _hasOtherWorkbookOverride(workbookToIgnore);
            }

            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        internal bool HasVisibleNonKernelWorkbook()
        {
            foreach (Excel.Workbook workbook in _application.Workbooks)
            {
                if (IsKernelWorkbook(workbook))
                {
                    continue;
                }

                if (workbook.Windows.Count > 0 && workbook.Windows.Cast<Excel.Window>().Any(window => window.Visible))
                {
                    return true;
                }
            }

            return false;
        }

        internal bool HasExplicitKernelStartupContext(Excel.Workbook startupWorkbook)
        {
            if (IsKernelWorkbook(startupWorkbook))
            {
                return true;
            }

            try
            {
                return IsKernelWorkbook(_application.ActiveWorkbook);
            }
            catch
            {
                return false;
            }
        }

        internal bool HasKernelWorkbookContext()
        {
            try
            {
                if (IsKernelWorkbook(_application.ActiveWorkbook))
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                return GetOpenKernelWorkbook() != null;
            }
            catch
            {
                return false;
            }
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
