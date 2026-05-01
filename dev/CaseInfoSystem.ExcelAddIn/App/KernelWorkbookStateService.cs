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
        private readonly KernelStartupContextInspector _kernelStartupContextInspector;
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
            _kernelStartupContextInspector = new KernelStartupContextInspector(_application, _excelInteropService, _kernelOpenWorkbookLocator);
            _hasOtherVisibleWorkbookOverride = hasOtherVisibleWorkbookOverride;
            _hasOtherWorkbookOverride = hasOtherWorkbookOverride;
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return WorkbookFileNameResolver.IsKernelWorkbookName(GetWorkbookNameCore(workbook));
        }

        internal bool ShouldShowHomeOnStartup(Excel.Workbook startupWorkbook = null)
        {
            KernelStartupContext context = _kernelStartupContextInspector.InspectForStartupDisplay(startupWorkbook);
            bool hasExplicitKernelStartupContext = context.StartupWorkbookIsKernel || context.StartupContextActiveWorkbookIsKernel;
            bool hasKernelWorkbookContext = hasExplicitKernelStartupContext
                && (context.KernelContextActiveWorkbookIsKernel || context.HasOpenKernelWorkbook);
            bool isStartupWorkbookKernel = hasExplicitKernelStartupContext && hasKernelWorkbookContext && context.StartupWorkbookIsKernel;
            bool hasVisibleNonKernelWorkbook = hasExplicitKernelStartupContext
                && hasKernelWorkbookContext
                && !isStartupWorkbookKernel
                && context.HasVisibleNonKernelWorkbook;
            return KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext,
                hasKernelWorkbookContext,
                isStartupWorkbookKernel,
                hasVisibleNonKernelWorkbook);
        }

        internal string DescribeStartupState()
        {
            KernelStartupContext context = _kernelStartupContextInspector.InspectForStartupDescription();
            return "activeWorkbook="
                + context.DescribeActiveWorkbookName
                + ", activeIsKernel="
                + context.DescribeActiveWorkbookIsKernel
                + ", hasOpenKernelWorkbook="
                + context.HasOpenKernelWorkbook
                + ", hasVisibleNonKernelWorkbook="
                + context.HasVisibleNonKernelWorkbook;
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
