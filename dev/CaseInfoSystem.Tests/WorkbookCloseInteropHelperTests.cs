using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class WorkbookCloseInteropHelperTests
    {
        [Fact]
        public void Close_PassesMissingOptionalArgumentsToWorkbookClose()
        {
            var application = new Excel.Application();
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\temp\Kernel.xlsx",
                Name = "Kernel.xlsx",
                Path = @"C:\temp"
            };

            application.Workbooks.Add(workbook);

            WorkbookCloseInteropHelper.Close(workbook);

            Assert.Equal(1, workbook.CloseCallCount);
            Assert.Same(Type.Missing, workbook.LastCloseSaveChangesArgument);
            Assert.Same(Type.Missing, workbook.LastCloseFilename);
            Assert.Same(Type.Missing, workbook.LastCloseRouteWorkbook);
        }

        [Fact]
        public void CloseWithoutSave_PassesFalseAndMissingOptionalArgumentsToWorkbookClose()
        {
            var application = new Excel.Application();
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\temp\Kernel.xlsx",
                Name = "Kernel.xlsx",
                Path = @"C:\temp"
            };

            application.Workbooks.Add(workbook);

            WorkbookCloseInteropHelper.CloseWithoutSave(workbook);

            Assert.Equal(1, workbook.CloseCallCount);
            Assert.False(workbook.LastCloseSaveChanges.GetValueOrDefault());
            Assert.Same(Type.Missing, workbook.LastCloseFilename);
            Assert.Same(Type.Missing, workbook.LastCloseRouteWorkbook);
        }

        [Fact]
        public void CloseWithoutSave_WhenWorkbookIsNull_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => WorkbookCloseInteropHelper.CloseWithoutSave(null));
        }
    }
}
