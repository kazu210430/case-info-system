using CaseInfoSystem.ExcelAddIn.App;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class WorkbookPromptSuppressionHelperTests
    {
        [Fact]
        public void MarkWorkbookSavedForPromptlessClose_WhenWorkbookWasDirty_MarksItCleanForClose()
        {
            var workbook = new Excel.Workbook
            {
                Saved = false
            };

            WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(workbook);

            Assert.True(workbook.Saved);
        }

        [Fact]
        public void MarkWorkbookSavedForPromptlessClose_WhenWorkbookWasAlreadyClean_LeavesItClean()
        {
            var workbook = new Excel.Workbook
            {
                Saved = true
            };

            WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(workbook);

            Assert.True(workbook.Saved);
        }

        [Fact]
        public void MarkWorkbookSavedForPromptlessClose_WhenWorkbookIsNull_DoesNothing()
        {
            WorkbookPromptSuppressionHelper.MarkWorkbookSavedForPromptlessClose(null);
        }
    }
}
