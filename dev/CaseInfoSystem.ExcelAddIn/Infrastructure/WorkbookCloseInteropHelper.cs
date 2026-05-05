using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class WorkbookCloseInteropHelper
    {
        internal static void Close(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
        }

        internal static void CloseWithoutSave(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            workbook.Close(false, Type.Missing, Type.Missing);
        }
    }
}
