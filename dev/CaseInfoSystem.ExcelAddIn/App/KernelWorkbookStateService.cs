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
        private readonly Logger _logger;
        private readonly KernelOpenWorkbookLocator _kernelOpenWorkbookLocator;
        private readonly Func<Excel.Workbook, bool> _hasOtherVisibleWorkbookOverride;
        private readonly Func<Excel.Workbook, bool> _hasOtherWorkbookOverride;

        internal KernelWorkbookStateService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            Logger logger,
            KernelOpenWorkbookLocator kernelOpenWorkbookLocator,
            Func<Excel.Workbook, bool> hasOtherVisibleWorkbookOverride = null,
            Func<Excel.Workbook, bool> hasOtherWorkbookOverride = null)
        {
            _application = application;
            _excelInteropService = excelInteropService;
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _kernelOpenWorkbookLocator = kernelOpenWorkbookLocator ?? throw new ArgumentNullException(nameof(kernelOpenWorkbookLocator));
            _hasOtherVisibleWorkbookOverride = hasOtherVisibleWorkbookOverride;
            _hasOtherWorkbookOverride = hasOtherWorkbookOverride;
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return WorkbookFileNameResolver.IsKernelWorkbookName(GetWorkbookNameCore(workbook));
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
                hasOpenKernelWorkbook = _kernelOpenWorkbookLocator.GetOpenKernelWorkbook() != null;
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
                return _kernelOpenWorkbookLocator.GetOpenKernelWorkbook() != null;
            }
            catch
            {
                return false;
            }
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
