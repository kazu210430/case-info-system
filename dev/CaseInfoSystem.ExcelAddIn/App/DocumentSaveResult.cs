namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentSaveResult
    {
        /// <summary>
        internal DocumentSaveResult(string savedPath, string finalPath, bool isLocalWorkCopy)
        {
            SavedPath = savedPath ?? string.Empty;
            FinalPath = finalPath ?? string.Empty;
            IsLocalWorkCopy = isLocalWorkCopy;
        }

        /// <summary>
        internal string SavedPath { get; }

        /// <summary>
        internal string FinalPath { get; }

        /// <summary>
        internal bool IsLocalWorkCopy { get; }
    }
}
