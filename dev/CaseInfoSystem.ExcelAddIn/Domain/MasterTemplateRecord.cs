namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class MasterTemplateRecord
    {
        internal string Key { get; set; } = string.Empty;

        internal string TemplateFileName { get; set; } = string.Empty;

        internal string DocumentName { get; set; } = string.Empty;

        internal long BackColor { get; set; }
    }
}
