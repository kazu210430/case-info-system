using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class AccountingSetContext
    {
        internal Excel.Workbook SourceWorkbook { get; set; }

        internal Excel.Worksheet SourceWorksheet { get; set; }

        internal string CustomerName { get; set; }

        internal string LawyerLinesText { get; set; }

        internal string SystemRoot { get; set; }

        internal string OutputFolderPath { get; set; }
    }
}
