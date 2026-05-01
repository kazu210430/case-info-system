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
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger));
            }

            if (kernelOpenWorkbookLocator == null)
            {
                throw new ArgumentNullException(nameof(kernelOpenWorkbookLocator));
            }

            _kernelStartupContextInspector = new KernelStartupContextInspector(
                _application,
                kernelOpenWorkbookLocator,
                HasVisibleNonKernelWorkbook);
            _hasOtherVisibleWorkbookOverride = hasOtherVisibleWorkbookOverride;
            _hasOtherWorkbookOverride = hasOtherWorkbookOverride;
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return WorkbookFileNameResolver.IsKernelWorkbookName(GetWorkbookNameCore(workbook));
        }

        internal bool ShouldShowHomeOnStartup(Excel.Workbook startupWorkbook = null)
        {
            KernelStartupContext context = _kernelStartupContextInspector.Inspect(startupWorkbook);
            bool hasExplicitKernelStartupContext = HasExplicitKernelStartupContext(context);
            bool hasKernelWorkbookContext = false;
            bool isStartupWorkbookKernel = false;
            bool hasVisibleNonKernelWorkbook = false;

            if (hasExplicitKernelStartupContext)
            {
                _kernelStartupContextInspector.PopulateOpenKernelWorkbookState(context);
                hasKernelWorkbookContext = HasKernelWorkbookContext(context);
                isStartupWorkbookKernel = hasKernelWorkbookContext && IsKernelWorkbook(context.StartupWorkbook);
                if (hasKernelWorkbookContext && !isStartupWorkbookKernel)
                {
                    _kernelStartupContextInspector.PopulateVisibleNonKernelWorkbookState(context);
                    hasVisibleNonKernelWorkbook = context.HasVisibleNonKernelWorkbook;
                }
            }

            return KernelWorkbookStartupDisplayPolicy.ShouldShowHomeOnStartup(
                hasExplicitKernelStartupContext,
                hasKernelWorkbookContext,
                isStartupWorkbookKernel,
                hasVisibleNonKernelWorkbook);
        }

        internal string DescribeStartupState()
        {
            KernelStartupContext context = _kernelStartupContextInspector.Inspect(startupWorkbook: null);
            string activeWorkbookName = context.ActiveWorkbookAccessFailed
                ? "(error)"
                : context.ActiveWorkbook == null
                    ? "(null)"
                    : GetWorkbookNameCore(context.ActiveWorkbook);
            bool activeIsKernel = context.ActiveWorkbook != null && IsKernelWorkbook(context.ActiveWorkbook);

            try
            {
                _kernelStartupContextInspector.PopulateOpenKernelWorkbookState(context);
                _kernelStartupContextInspector.PopulateVisibleNonKernelWorkbookState(context);
            }
            catch
            {
            }

            return "activeWorkbook="
                + activeWorkbookName
                + ", activeIsKernel="
                + activeIsKernel
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

        private bool HasExplicitKernelStartupContext(KernelStartupContext context)
        {
            if (context == null)
            {
                return false;
            }

            if (IsKernelWorkbook(context.StartupWorkbook))
            {
                return true;
            }

            return IsKernelWorkbook(context.ActiveWorkbook);
        }

        private bool HasKernelWorkbookContext(KernelStartupContext context)
        {
            return context != null
                && (IsKernelWorkbook(context.ActiveWorkbook) || context.HasOpenKernelWorkbook);
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
