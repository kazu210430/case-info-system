using System;
using System.Collections.Generic;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class CaseContextFactory
    {
        private readonly ExcelInteropService _excelInteropService;
        private readonly CaseDataSnapshotFactory _caseDataSnapshotFactory;
        private readonly Logger _logger;

        /// <summary>
        internal CaseContextFactory(
            ExcelInteropService excelInteropService,
            CaseDataSnapshotFactory caseDataSnapshotFactory,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _caseDataSnapshotFactory = caseDataSnapshotFactory ?? throw new ArgumentNullException(nameof(caseDataSnapshotFactory));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal CaseContext CreateForCaseListRegistration(Excel.Workbook caseWorkbook)
        {
            if (caseWorkbook == null)
            {
                return null;
            }

            string rowText = _excelInteropService.TryGetDocumentProperty(caseWorkbook, "CASELIST_ROW");
            int registeredRow;
            if (!int.TryParse(rowText, out registeredRow) || registeredRow <= 0)
            {
                _logger.Info("CASELIST_ROW could not be resolved after registration.");
                return null;
            }

            Excel.Workbook kernelWorkbook = _excelInteropService.FindKernelWorkbook(caseWorkbook);
            Excel.Worksheet caseListWorksheet = _excelInteropService.FindCaseListWorksheet(kernelWorkbook);
            if (kernelWorkbook == null || caseListWorksheet == null)
            {
                _logger.Info("Kernel workbook or case list sheet was not available for row normalization.");
                return null;
            }

            return new CaseContext
            {
                CaseWorkbook = caseWorkbook,
                KernelWorkbook = kernelWorkbook,
                CaseListWorksheet = caseListWorksheet,
                RegisteredRow = registeredRow,
                SystemRoot = _excelInteropService.TryGetDocumentProperty(caseWorkbook, "SYSTEM_ROOT"),
                WorkbookName = _excelInteropService.GetWorkbookName(caseWorkbook),
                WorkbookPath = _excelInteropService.GetWorkbookPath(caseWorkbook)
            };
        }

        /// <summary>
        internal CaseContext CreateForDocumentCreate(Excel.Workbook caseWorkbook)
        {
            if (caseWorkbook == null)
            {
                return null;
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            Excel.Worksheet homeWorksheet = _excelInteropService.FindWorksheetByCodeName(caseWorkbook, "shHOME");
            CaseDataSnapshot caseDataSnapshot = _caseDataSnapshotFactory.Create(caseWorkbook);
            IReadOnlyDictionary<string, string> caseValues = caseDataSnapshot == null
                ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                : caseDataSnapshot.Values;

            CaseContext context = new CaseContext
            {
                CaseWorkbook = caseWorkbook,
                HomeWorksheet = homeWorksheet,
                CaseValues = caseValues,
                CustomerName = caseDataSnapshot == null ? string.Empty : caseDataSnapshot.CustomerName,
                WorkbookName = _excelInteropService.GetWorkbookName(caseWorkbook),
                WorkbookPath = _excelInteropService.GetWorkbookPath(caseWorkbook),
                HomeSheetName = homeWorksheet == null ? string.Empty : homeWorksheet.Name,
                SystemRoot = _excelInteropService.TryGetDocumentProperty(caseWorkbook, "SYSTEM_ROOT")
            };
            _logger.Debug(
                "CaseContextFactory.CreateForDocumentCreate",
                "Completed elapsed=" + FormatElapsedSeconds(stopwatch.Elapsed)
                + " workbook=" + (context.WorkbookName ?? string.Empty)
                + " caseValueCount=" + context.CaseValues.Count.ToString()
                + " customerNameLength=" + (context.CustomerName ?? string.Empty).Length.ToString());
            return context;
        }

        private static string FormatElapsedSeconds(TimeSpan elapsed)
        {
            return elapsed.TotalSeconds.ToString("0.000");
        }
    }
}
