using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class WorkbookContextSeed
    {
        internal WorkbookContextSeed(
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
