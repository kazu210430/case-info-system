using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelStartupContext
    {
        internal Excel.Workbook StartupWorkbook { get; set; }

        internal Excel.Workbook ActiveWorkbook { get; set; }

        internal bool HasOpenKernelWorkbook { get; set; }

        internal bool HasVisibleNonKernelWorkbook { get; set; }

        internal bool ActiveWorkbookAccessFailed { get; set; }
    }
}
