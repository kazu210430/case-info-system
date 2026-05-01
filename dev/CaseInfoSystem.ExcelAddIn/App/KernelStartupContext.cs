using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelStartupContext
    {
        internal Excel.Workbook StartupWorkbook { get; set; }

        internal bool StartupWorkbookIsKernel { get; set; }

        internal bool StartupContextActiveWorkbookIsKernel { get; set; }

        internal bool StartupContextActiveWorkbookReadFailed { get; set; }

        internal bool KernelContextActiveWorkbookIsKernel { get; set; }

        internal bool KernelContextActiveWorkbookReadFailed { get; set; }

        internal bool HasOpenKernelWorkbook { get; set; }

        internal bool HasVisibleNonKernelWorkbook { get; set; }

        internal string DescribeActiveWorkbookName { get; set; } = "(null)";

        internal bool DescribeActiveWorkbookIsKernel { get; set; }

        internal bool DescribeActiveWorkbookReadFailed { get; set; }
    }
}
