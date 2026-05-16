using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class KernelCaseCreationResult
    {
        internal bool Success { get; set; }

        internal KernelCaseCreationMode Mode { get; set; }

        internal string CaseFolderPath { get; set; }

        internal string CaseWorkbookPath { get; set; }

        internal Excel.Workbook CreatedWorkbook { get; set; }

        internal string UserMessage { get; set; }

        internal bool ShouldCloseKernelHome { get; set; }

        internal CasePresentationOutcome PresentationOutcome { get; set; }

        internal string PresentationOutcomeReason { get; set; }
    }

    internal enum CasePresentationOutcome
    {
        NotStarted = 0,
        Completed = 1,
        Degraded = 2,
        Failed = 3
    }
}
