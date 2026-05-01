using System.Linq;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelStartupContextInspector
    {
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly KernelOpenWorkbookLocator _kernelOpenWorkbookLocator;

        internal KernelStartupContextInspector(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            KernelOpenWorkbookLocator kernelOpenWorkbookLocator)
        {
            _application = application;
            _excelInteropService = excelInteropService;
            _kernelOpenWorkbookLocator = kernelOpenWorkbookLocator;
        }

        internal KernelStartupContext InspectForStartupDisplay(Excel.Workbook startupWorkbook)
        {
            var context = new KernelStartupContext
            {
                StartupWorkbook = startupWorkbook,
                StartupWorkbookIsKernel = IsKernelWorkbook(startupWorkbook)
            };

            if (!context.StartupWorkbookIsKernel)
            {
                try
                {
                    context.StartupContextActiveWorkbookIsKernel = IsKernelWorkbook(_application.ActiveWorkbook);
                }
                catch
                {
                    context.StartupContextActiveWorkbookReadFailed = true;
                }
            }

            bool hasExplicitKernelStartupContext = context.StartupWorkbookIsKernel || context.StartupContextActiveWorkbookIsKernel;
            if (!hasExplicitKernelStartupContext)
            {
                return context;
            }

            try
            {
                context.KernelContextActiveWorkbookIsKernel = IsKernelWorkbook(_application.ActiveWorkbook);
            }
            catch
            {
                context.KernelContextActiveWorkbookReadFailed = true;
            }

            if (!context.KernelContextActiveWorkbookIsKernel)
            {
                try
                {
                    context.HasOpenKernelWorkbook = _kernelOpenWorkbookLocator.GetOpenKernelWorkbook() != null;
                }
                catch
                {
                }
            }

            bool hasKernelWorkbookContext = context.KernelContextActiveWorkbookIsKernel || context.HasOpenKernelWorkbook;
            if (!hasKernelWorkbookContext || context.StartupWorkbookIsKernel)
            {
                return context;
            }

            context.HasVisibleNonKernelWorkbook = HasVisibleNonKernelWorkbook();
            return context;
        }

        internal KernelStartupContext InspectForStartupDescription()
        {
            var context = new KernelStartupContext();

            try
            {
                Excel.Workbook activeWorkbook = _application.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    context.DescribeActiveWorkbookName = GetWorkbookNameCore(activeWorkbook);
                    context.DescribeActiveWorkbookIsKernel = IsKernelWorkbook(activeWorkbook);
                }
            }
            catch
            {
                context.DescribeActiveWorkbookName = "(error)";
                context.DescribeActiveWorkbookReadFailed = true;
            }

            try
            {
                context.HasOpenKernelWorkbook = _kernelOpenWorkbookLocator.GetOpenKernelWorkbook() != null;
                context.HasVisibleNonKernelWorkbook = HasVisibleNonKernelWorkbook();
            }
            catch
            {
            }

            return context;
        }

        private bool HasVisibleNonKernelWorkbook()
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
