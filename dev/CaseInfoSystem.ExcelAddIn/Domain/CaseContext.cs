using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class CaseContext
    {
        internal Excel.Workbook CaseWorkbook { get; set; }

        internal Excel.Workbook KernelWorkbook { get; set; }

        internal Excel.Worksheet CaseListWorksheet { get; set; }

        internal Excel.Worksheet HomeWorksheet { get; set; }

        internal int RegisteredRow { get; set; }

        internal string SystemRoot { get; set; }

        internal string CustomerName { get; set; }

        internal string WorkbookPath { get; set; }

        internal string WorkbookName { get; set; }

        internal string HomeSheetName { get; set; }

        internal IReadOnlyDictionary<string, string> CaseValues { get; set; }
    }
}
