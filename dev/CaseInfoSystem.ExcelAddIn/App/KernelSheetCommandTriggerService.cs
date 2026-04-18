using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelSheetCommandTriggerService
    {
        private readonly KernelCommandService _kernelCommandService;
        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly ExcelInteropService _excelInteropService;
        private readonly Excel.Application _application;
        private readonly string _kernelSheetCommandSheetCodeName;
        private readonly string _kernelSheetCommandCellAddress;
        private readonly Action<Excel.Range> _clearKernelSheetCommandCell;
        private readonly Action<object> _releaseComObject;
        private readonly Logger _logger;

        internal KernelSheetCommandTriggerService(
            KernelCommandService kernelCommandService,
            KernelWorkbookService kernelWorkbookService,
            ExcelInteropService excelInteropService,
            Excel.Application application,
            string kernelSheetCommandSheetCodeName,
            string kernelSheetCommandCellAddress,
            Action<Excel.Range> clearKernelSheetCommandCell,
            Action<object> releaseComObject,
            Logger logger)
        {
            _kernelCommandService = kernelCommandService;
            _kernelWorkbookService = kernelWorkbookService;
            _excelInteropService = excelInteropService;
            _application = application;
            _kernelSheetCommandSheetCodeName = kernelSheetCommandSheetCodeName;
            _kernelSheetCommandCellAddress = kernelSheetCommandCellAddress;
            _clearKernelSheetCommandCell = clearKernelSheetCommandCell;
            _releaseComObject = releaseComObject;
            _logger = logger;
        }

        internal void Handle(Excel.Worksheet worksheet, Excel.Range target)
        {
            if (_kernelCommandService == null || _kernelWorkbookService == null || worksheet == null || target == null)
            {
                return;
            }

            Excel.Range commandCell = null;
            Excel.Range intersection = null;
            try
            {
                Excel.Workbook workbook = worksheet.Parent as Excel.Workbook;
                if (workbook == null || !_kernelWorkbookService.IsKernelWorkbook(workbook))
                {
                    return;
                }

                if (!string.Equals(worksheet.CodeName, _kernelSheetCommandSheetCodeName, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                commandCell = worksheet.Range[_kernelSheetCommandCellAddress];
                intersection = _application.Intersect(target, commandCell);
                if (intersection == null)
                {
                    return;
                }

                string commandText = (Convert.ToString(commandCell.Value2) ?? string.Empty).Trim();
                if (commandText.Length == 0)
                {
                    return;
                }

                _logger?.Info("Kernel sheet command detected. command=" + commandText + ", workbook=" + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));
                _clearKernelSheetCommandCell(commandCell);
                _kernelCommandService.ExecuteSheetCommand(commandText);
            }
            catch (Exception ex)
            {
                _logger?.Error("HandleKernelSheetCommand failed.", ex);
            }
            finally
            {
                _releaseComObject(intersection);
                _releaseComObject(commandCell);
            }
        }
    }
}
