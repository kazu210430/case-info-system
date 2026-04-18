namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class KernelCaseCreationPlan
    {
        internal KernelCaseCreationMode Mode { get; set; }

        internal string CustomerName { get; set; }

        internal string SystemRoot { get; set; }

        internal string BaseWorkbookPath { get; set; }

        internal string CaseFolderPath { get; set; }

        internal string CaseWorkbookPath { get; set; }

        internal string NameRuleA { get; set; }

        internal string NameRuleB { get; set; }
    }
}
