using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class SheetEventCoordinator
    {
        private readonly Logger _logger;
        private readonly KernelSheetCommandTriggerService _kernelSheetCommandTriggerService;
        private readonly CaseWorkbookLifecycleService _caseWorkbookLifecycleService;
        private readonly AccountingWorkbookLifecycleService _accountingWorkbookLifecycleService;
        private readonly AccountingSheetControlService _accountingSheetControlService;
        private readonly Action<string, Excel.Workbook, Excel.Window> _refreshTaskPane;

        internal SheetEventCoordinator(
            Logger logger,
            KernelSheetCommandTriggerService kernelSheetCommandTriggerService,
            CaseWorkbookLifecycleService caseWorkbookLifecycleService,
            AccountingWorkbookLifecycleService accountingWorkbookLifecycleService,
            AccountingSheetControlService accountingSheetControlService,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPane)
        {
            _logger = logger;
            _kernelSheetCommandTriggerService = kernelSheetCommandTriggerService;
            _caseWorkbookLifecycleService = caseWorkbookLifecycleService;
            _accountingWorkbookLifecycleService = accountingWorkbookLifecycleService;
            _accountingSheetControlService = accountingSheetControlService;
            _refreshTaskPane = refreshTaskPane;
        }

        internal void OnSheetActivate(object sheetObject)
        {
            _accountingWorkbookLifecycleService?.HandleSheetActivated(sheetObject);
            _accountingSheetControlService?.HandleSheetActivated(sheetObject);
            _caseWorkbookLifecycleService?.HandleSheetActivated(sheetObject);
            _refreshTaskPane?.Invoke("SheetActivate", null, null);
        }

        internal void OnSheetChange(object sheetObject, Excel.Range target)
        {
            string sheetName = SafeSheetName(sheetObject);
            string targetAddress = SafeRangeAddress(target);
            _logger?.Debug("Application_SheetChange", "fired. sheet=" + sheetName + ", target=" + targetAddress);

            Excel.Worksheet worksheet = sheetObject as Excel.Worksheet;
            _kernelSheetCommandTriggerService?.Handle(worksheet, target);
            _caseWorkbookLifecycleService?.HandleSheetChanged(worksheet == null ? null : worksheet.Parent as Excel.Workbook);
            _accountingSheetControlService?.HandleSheetChange(sheetObject, target);
        }

        internal void OnSheetSelectionChange(object sheetObject, Excel.Range target)
        {
            string sheetName = SafeSheetName(sheetObject);
            string targetAddress = SafeRangeAddress(target);
            _logger?.Debug("Application_SheetSelectionChange", "fired. sheet=" + sheetName + ", target=" + targetAddress);
            _accountingSheetControlService?.HandleSheetSelectionChange(sheetObject, target);
        }

        internal void OnAfterCalculate(Excel.Application application)
        {
            _logger?.Debug("Application_AfterCalculate", "fired.");
            _accountingSheetControlService?.HandleAfterCalculate(application);
            _logger?.Debug("Application_AfterCalculate", "after AccountingSheetControlService.HandleAfterCalculate returned.");
        }

        private static string SafeSheetName(object sheetObject)
        {
            try
            {
                var worksheet = sheetObject as Excel.Worksheet;
                return worksheet == null ? string.Empty : worksheet.CodeName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeRangeAddress(Excel.Range range)
        {
            try
            {
                return range == null ? string.Empty : Convert.ToString(range.Address[false, false]) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
