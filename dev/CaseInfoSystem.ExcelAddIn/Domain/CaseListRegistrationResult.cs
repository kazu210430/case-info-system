namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class CaseListRegistrationResult
    {
        internal bool Success { get; set; }

        internal int RegisteredRow { get; set; }

        internal string Message { get; set; } = string.Empty;
    }
}
