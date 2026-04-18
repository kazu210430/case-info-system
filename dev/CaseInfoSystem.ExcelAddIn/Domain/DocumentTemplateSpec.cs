namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class DocumentTemplateSpec
    {
        internal string Key { get; set; }

        internal string DocumentName { get; set; }

        internal string TemplateFileName { get; set; }

        internal string TemplatePath { get; set; }

        internal string ActionKind { get; set; }

        internal DocumentTemplateResolutionSource ResolutionSource { get; set; }
    }
}
