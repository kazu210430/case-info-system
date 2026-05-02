using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class CaseFolderOpenService
    {
        private const string CreatedCaseFolderOfferOpenReason = "CaseWorkbookLifecycleService.PostCloseCreatedCaseFolderOffer";

        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly FolderWindowService _folderWindowService;

        internal CaseFolderOpenService(
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            FolderWindowService folderWindowService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _folderWindowService = folderWindowService ?? throw new ArgumentNullException(nameof(folderWindowService));
        }

        internal string ResolveContainingFolder(Excel.Workbook workbook)
        {
            string folderPath = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(workbook));
            if (folderPath.Length == 0 || !_pathCompatibilityService.DirectoryExistsSafe(folderPath))
            {
                return string.Empty;
            }

            return folderPath;
        }

        internal bool DirectoryExistsSafe(string folderPath)
        {
            return _pathCompatibilityService.DirectoryExistsSafe(folderPath);
        }

        internal void OpenCreatedCaseFolder(string folderPath)
        {
            _folderWindowService.OpenFolder(folderPath, CreatedCaseFolderOfferOpenReason);
        }
    }
}
