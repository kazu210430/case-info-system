using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class WorkbookSessionService
    {
        private readonly NavigationService _navigationService;
        private readonly TransientPaneSuppressionService _transientPaneSuppressionService;
        private readonly Logger _logger;

        /// <summary>
        internal WorkbookSessionService(
            NavigationService navigationService,
            TransientPaneSuppressionService transientPaneSuppressionService,
            Logger logger)
        {
            _navigationService = navigationService ?? throw new ArgumentNullException(nameof(navigationService));
            _transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException(nameof(transientPaneSuppressionService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal WorkbookContext ResolveActiveContext(string reason)
        {
            WorkbookContext context = _navigationService.ResolveActiveContext();
            _navigationService.TraceContext(context, reason);
            return context;
        }

        /// <summary>
        internal WorkbookContext ResolveContext(Excel.Workbook workbook, Excel.Window window, string reason)
        {
            WorkbookContext context = _navigationService.CreateContext(workbook, window);
            _navigationService.TraceContext(context, reason);
            return context;
        }

        /// <summary>
        internal bool ShouldHandleContext(WorkbookContext context)
        {
            if (context == null)
            {
                _logger.Warn("ShouldHandleContext skipped because context was null.");
                return false;
            }

            bool isSuppressed = context.Workbook != null
                ? _transientPaneSuppressionService.IsSuppressed(context.Workbook)
                : _transientPaneSuppressionService.IsSuppressedPath(context.WorkbookFullName);

            if (isSuppressed)
            {
                _logger.Info("ShouldHandleContext=False because workbook is transiently suppressed. workbook=" + context.WorkbookFullName);
                return false;
            }

            bool handled =
                context.Role == WorkbookRole.Kernel
                || context.Role == WorkbookRole.Case
                || context.Role == WorkbookRole.Accounting;
            if (!handled)
            {
                _logger.Info("ShouldHandleContext=False, workbook=" + context.WorkbookFullName);
            }

            return handled;
        }
    }
}
