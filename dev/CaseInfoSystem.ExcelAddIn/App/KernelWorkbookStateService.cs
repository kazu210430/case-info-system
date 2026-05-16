using System;
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

            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = _application == null ? null : _application.Workbooks;
                int workbookCount = workbooks == null ? 0 : workbooks.Count;
                for (int index = 1; index <= workbookCount; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                        {
                            continue;
                        }

                        if (HasVisibleWindow(workbook))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        if (!ReferenceEquals(workbook, workbookToIgnore))
                        {
                            ComObjectReleaseService.Release(workbook, _logger);
                        }
                    }
                }
            }
            finally
            {
                ComObjectReleaseService.Release(workbooks, _logger);
            }

            return false;
        }

        internal bool HasOtherWorkbook(Excel.Workbook workbookToIgnore)
        {
            if (_hasOtherWorkbookOverride != null)
            {
                return _hasOtherWorkbookOverride(workbookToIgnore);
            }

            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = _application == null ? null : _application.Workbooks;
                int workbookCount = workbooks == null ? 0 : workbooks.Count;
                for (int index = 1; index <= workbookCount; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        if (workbookToIgnore != null && ReferenceEquals(workbook, workbookToIgnore))
                        {
                            continue;
                        }

                        return true;
                    }
                    finally
                    {
                        if (!ReferenceEquals(workbook, workbookToIgnore))
                        {
                            ComObjectReleaseService.Release(workbook, _logger);
                        }
                    }
                }
            }
            finally
            {
                ComObjectReleaseService.Release(workbooks, _logger);
            }

            return false;
        }

        internal bool HasVisibleNonKernelWorkbook()
        {
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = _application == null ? null : _application.Workbooks;
                int workbookCount = workbooks == null ? 0 : workbooks.Count;
                for (int index = 1; index <= workbookCount; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        if (IsKernelWorkbook(workbook))
                        {
                            continue;
                        }

                        if (HasVisibleWindow(workbook))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ComObjectReleaseService.Release(workbook, _logger);
                    }
                }
            }
            finally
            {
                ComObjectReleaseService.Release(workbooks, _logger);
            }

            return false;
        }

        private bool HasVisibleWindow(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            Excel.Windows windows = null;
            try
            {
                windows = workbook.Windows;
                int windowCount = windows == null ? 0 : windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = windows[index];
                        if (window != null && window.Visible)
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ComObjectReleaseService.Release(window, _logger);
                    }
                }
            }
            finally
            {
                ComObjectReleaseService.Release(windows, _logger);
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
