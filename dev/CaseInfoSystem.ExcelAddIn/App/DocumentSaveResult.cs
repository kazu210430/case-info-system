namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentSaveResult
    {
        /// <summary>
        internal DocumentSaveResult(string savedPath, string finalPath, bool isLocalWorkCopy, object activeDocument = null)
        {
            SavedPath = savedPath ?? string.Empty;
            FinalPath = finalPath ?? string.Empty;
            IsLocalWorkCopy = isLocalWorkCopy;
            ActiveDocument = activeDocument;
        }

        /// <summary>
        internal string SavedPath { get; }

        /// <summary>
        internal string FinalPath { get; }

        /// <summary>
        internal bool IsLocalWorkCopy { get; }

        /// <summary>
        internal object ActiveDocument { get; }
    }
}
