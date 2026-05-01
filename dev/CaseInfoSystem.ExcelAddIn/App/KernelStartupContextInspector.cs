using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelStartupContextInspector
    {
        private readonly Excel.Application _application;
        private readonly KernelOpenWorkbookLocator _kernelOpenWorkbookLocator;
        private readonly Func<bool> _hasVisibleNonKernelWorkbook;

        internal KernelStartupContextInspector(
            Excel.Application application,
            KernelOpenWorkbookLocator kernelOpenWorkbookLocator,
            Func<bool> hasVisibleNonKernelWorkbook)
        {
            _application = application;
            _kernelOpenWorkbookLocator = kernelOpenWorkbookLocator ?? throw new ArgumentNullException(nameof(kernelOpenWorkbookLocator));
            _hasVisibleNonKernelWorkbook = hasVisibleNonKernelWorkbook ?? throw new ArgumentNullException(nameof(hasVisibleNonKernelWorkbook));
        }

        internal KernelStartupContext Inspect(Excel.Workbook startupWorkbook)
        {
            var context = new KernelStartupContext
            {
                StartupWorkbook = startupWorkbook
            };

            if (_application == null)
            {
                return context;
            }

            try
            {
                context.ActiveWorkbook = _application.ActiveWorkbook;
            }
            catch
            {
                context.ActiveWorkbookAccessFailed = true;
            }

            return context;
        }

        internal void PopulateOpenKernelWorkbookState(KernelStartupContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            try
            {
                context.HasOpenKernelWorkbook = _kernelOpenWorkbookLocator.GetOpenKernelWorkbook() != null;
            }
            catch
            {
                context.HasOpenKernelWorkbook = false;
            }
        }

        internal void PopulateVisibleNonKernelWorkbookState(KernelStartupContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            context.HasVisibleNonKernelWorkbook = _hasVisibleNonKernelWorkbook();
        }
    }
}
