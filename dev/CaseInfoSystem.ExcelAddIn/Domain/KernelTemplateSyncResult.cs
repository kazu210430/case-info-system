namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class KernelTemplateSyncResult
    {
        internal bool Success { get; set; }

        internal int UpdatedCount { get; set; }

        internal int DetectedCount { get; set; }

        internal int MasterVersion { get; set; }

        internal string TemplateDirectory { get; set; } = string.Empty;

        internal string DuplicateInfo { get; set; } = string.Empty;

        internal string BaseSyncError { get; set; } = string.Empty;

        internal string Message { get; set; } = string.Empty;
    }
}
