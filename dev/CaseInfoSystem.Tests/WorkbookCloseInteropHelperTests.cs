using System;
using System.Collections.Generic;
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

        [Fact]
        public void CloseOwnedWorkbookWithoutSave_LogsRouteAndCloseResult()
        {
            var logs = new List<string>();
            var logger = new Logger(logs.Add);
            var application = new Excel.Application();
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\temp\Owned.xlsx",
                Name = "Owned.xlsx",
                Path = @"C:\temp"
            };

            application.Workbooks.Add(workbook);

            WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave(workbook, logger, "unit-owned-close");

            Assert.Equal(1, workbook.CloseCallCount);
            Assert.False(workbook.LastCloseSaveChanges.GetValueOrDefault());
            Assert.Contains(logs, message => ContainsOrdinal(message, "Workbook close starting")
                && ContainsOrdinal(message, "route=unit-owned-close")
                && ContainsOrdinal(message, "workbookName=Owned.xlsx")
                && ContainsOrdinal(message, "saveChanges=false"));
            Assert.Contains(logs, message => ContainsOrdinal(message, "Workbook close completed")
                && ContainsOrdinal(message, "closeSucceeded=true")
                && ContainsOrdinal(message, "closeHelperContract=caller-does-not-read-closed-workbook"));
        }

        [Fact]
        public void CloseReadOnlyWithoutSave_WhenCloseThrows_LogsFailureAndRethrows()
        {
            var logs = new List<string>();
            var logger = new Logger(logs.Add);
            var workbook = new Excel.Workbook
            {
                FullName = @"C:\temp\ReadOnly.xlsx",
                Name = "ReadOnly.xlsx",
                Path = @"C:\temp",
                CloseBehavior = () => throw new InvalidOperationException("close failed")
            };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => WorkbookCloseInteropHelper.CloseReadOnlyWithoutSave(workbook, logger, "unit-readonly-close"));

            Assert.Equal("close failed", exception.Message);
            Assert.Equal(1, workbook.CloseCallCount);
            Assert.False(workbook.LastCloseSaveChanges.GetValueOrDefault());
            Assert.Contains(logs, message => ContainsOrdinal(message, "Workbook close failed")
                && ContainsOrdinal(message, "route=unit-readonly-close")
                && ContainsOrdinal(message, "workbookName=ReadOnly.xlsx")
                && ContainsOrdinal(message, "saveChanges=false")
                && ContainsOrdinal(message, "InvalidOperationException"));
        }

        private static bool ContainsOrdinal(string value, string expected)
        {
            return (value ?? string.Empty).IndexOf(expected ?? string.Empty, StringComparison.Ordinal) >= 0;
        }
    }
}
