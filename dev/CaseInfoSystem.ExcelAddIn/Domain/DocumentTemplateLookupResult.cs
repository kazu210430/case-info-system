namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class DocumentTemplateLookupResult
    {
        internal string Key { get; set; } = string.Empty;

        internal string DocumentName { get; set; } = string.Empty;

        internal string TemplateFileName { get; set; } = string.Empty;

        internal DocumentTemplateResolutionSource ResolutionSource { get; set; }
    }
}
