using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal interface IMasterTemplateSheetReader
    {
        MasterTemplateSheetData Read(Excel.Worksheet worksheet);
    }

    internal sealed class MasterTemplateSheetReaderAdapter : IMasterTemplateSheetReader
    {
        internal MasterTemplateSheetReaderAdapter()
        {
        }

        public MasterTemplateSheetData Read(Excel.Worksheet worksheet)
        {
            return MasterTemplateSheetReader.Read(worksheet);
        }
    }
}
