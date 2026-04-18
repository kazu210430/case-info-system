namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class KernelCaseCreationRequest
    {
        internal string CustomerName { get; set; }

        internal KernelCaseCreationMode Mode { get; set; }

        internal string DefaultRoot { get; set; }

        internal string SelectedFolderPath { get; set; }
    }
}
