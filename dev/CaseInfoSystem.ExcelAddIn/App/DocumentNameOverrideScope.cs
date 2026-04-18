using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentNameOverrideScope : IDisposable
    {
        private const string SuppressCaseRevealPropertyName = "TASKPANE_SUPPRESS_CASE_REVEAL";
        private const string OverrideEnabledPropertyName = "TASKPANE_DOC_NAME_OVERRIDE_ENABLED";
        private const string OverrideValuePropertyName = "TASKPANE_DOC_NAME_OVERRIDE";

        private readonly ExcelInteropService _excelInteropService;
        private readonly Excel.Workbook _workbook;
        private readonly Logger _logger;
        private bool _disposed;

        /// <summary>
        internal DocumentNameOverrideScope(
            ExcelInteropService excelInteropService,
            Excel.Workbook workbook,
            Logger logger,
            string finalDocumentName)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            string documentName = finalDocumentName ?? string.Empty;
            _excelInteropService.SetDocumentProperty(_workbook, SuppressCaseRevealPropertyName, "1");
            _excelInteropService.SetDocumentProperty(_workbook, OverrideEnabledPropertyName, "1");
            _excelInteropService.SetDocumentProperty(_workbook, OverrideValuePropertyName, documentName);
        }

        /// <summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            try
            {
                _excelInteropService.SetDocumentProperty(_workbook, OverrideEnabledPropertyName, string.Empty);
                _excelInteropService.SetDocumentProperty(_workbook, OverrideValuePropertyName, string.Empty);
                _excelInteropService.SetDocumentProperty(_workbook, SuppressCaseRevealPropertyName, string.Empty);
            }
            catch (Exception ex)
            {
                _logger.Debug(nameof(DocumentNameOverrideScope), "Document name override cleanup failed but dispose continues. message=" + ex.Message);
            }
        }
    }
}
