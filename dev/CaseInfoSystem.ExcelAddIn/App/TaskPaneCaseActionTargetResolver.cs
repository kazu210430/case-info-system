using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneCaseActionTargetResolver
    {
        private readonly ExcelInteropService _excelInteropService;
        private readonly Logger _logger;
        private readonly Func<string, TaskPaneHost> _resolveHost;

        internal TaskPaneCaseActionTargetResolver(
            ExcelInteropService excelInteropService,
            Logger logger,
            Func<string, TaskPaneHost> resolveHost)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _resolveHost = resolveHost ?? throw new ArgumentNullException(nameof(resolveHost));
        }

        internal bool TryResolve(string windowKey, out TaskPaneHost host, out Excel.Workbook workbook)
        {
            host = null;
            workbook = null;

            if (string.IsNullOrWhiteSpace(windowKey))
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because host identity was not available.");
                return false;
            }

            host = _resolveHost(windowKey);
            if (host == null)
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because host was not found. windowKey=" + windowKey);
                return false;
            }

            workbook = _excelInteropService.FindOpenWorkbook(host.WorkbookFullName);
            if (workbook == null)
            {
                _logger.Warn("CaseControl_ActionInvoked skipped because workbook was not found. windowKey=" + windowKey);
                return false;
            }

            return true;
        }
    }
}
