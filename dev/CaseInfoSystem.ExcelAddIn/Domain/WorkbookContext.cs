using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class WorkbookContext
    {
        /// <summary>
        internal WorkbookContext(
            Excel.Workbook workbook,
            Excel.Window window,
            WorkbookRole role,
            string systemRoot,
            string workbookFullName,
            string activeSheetCodeName)
        {
            Workbook = workbook;
            Window = window;
            Role = role;
            SystemRoot = systemRoot ?? string.Empty;
            WorkbookFullName = workbookFullName ?? string.Empty;
            ActiveSheetCodeName = activeSheetCodeName ?? string.Empty;
        }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window Window { get; }

        internal WorkbookRole Role { get; }

        internal string SystemRoot { get; }

        internal string WorkbookFullName { get; }

        internal string ActiveSheetCodeName { get; }
    }
}
