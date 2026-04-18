using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class WorkbookPromptSuppressionHelper
    {
        // 「保存しないで閉じる」を明示的に選んだ経路や transient read-only 終了で、
        // Excel の確認ダイアログを出さずに close を完了させるために使う。
        internal static void MarkWorkbookSavedForPromptlessClose(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            workbook.Saved = true;
        }
    }
}
