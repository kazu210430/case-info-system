using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class KernelCasePathService
    {
        private readonly PathCompatibilityService _pathCompatibilityService;

        internal KernelCasePathService(PathCompatibilityService pathCompatibilityService)
        {
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
        }

        internal string ResolveSystemRoot(Excel.Workbook kernelWorkbook)
        {
            if (kernelWorkbook == null)
            {
                return string.Empty;
            }

            return _pathCompatibilityService.NormalizePath(kernelWorkbook.Path);
        }

        internal string ResolveBaseWorkbookPath(string systemRoot)
        {
            return WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        internal string ResolveCaseWorkbookExtension(string baseWorkbookPath)
        {
            return WorkbookFileNameResolver.GetWorkbookExtensionOrDefault(baseWorkbookPath);
        }

        internal string ResolveCaseFolderPath(KernelCaseCreationRequest request, string folderName)
        {
            if (request == null)
            {
                return string.Empty;
            }

            if (request.Mode == KernelCaseCreationMode.NewCaseDefault)
            {
                string folderPath = _pathCompatibilityService.CombinePath(request.DefaultRoot, folderName);
                return _pathCompatibilityService.EnsureUniqueDirectoryPathStandard(folderPath);
            }

            return _pathCompatibilityService.NormalizePath(request.SelectedFolderPath);
        }

        internal string BuildCaseWorkbookPath(string folderPath, string caseWorkbookName)
        {
            string workbookPath = _pathCompatibilityService.CombinePath(folderPath, caseWorkbookName);
            return _pathCompatibilityService.EnsureUniquePathStandard(workbookPath);
        }

        internal bool IsUnderSyncRoot(string path)
        {
            return _pathCompatibilityService.IsUnderSyncRoot(path);
        }

        internal string BuildLocalWorkingCaseWorkbookPath(string finalCaseWorkbookPath)
        {
            string normalizedFinalPath = _pathCompatibilityService.NormalizePath(finalCaseWorkbookPath);
            if (string.IsNullOrWhiteSpace(normalizedFinalPath))
            {
                return string.Empty;
            }

            string tempFolder = _pathCompatibilityService.GetLocalTempWorkFolder("CaseWorkbookTemp");
            if (string.IsNullOrWhiteSpace(tempFolder))
            {
                return string.Empty;
            }

            string fileName = _pathCompatibilityService.GetFileNameFromPath(normalizedFinalPath);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return string.Empty;
            }

            int dotPosition = fileName.LastIndexOf('.');
            string baseName = dotPosition > 1 ? fileName.Substring(0, dotPosition) : fileName;
            string extension = dotPosition > 1 ? fileName.Substring(dotPosition) : ".xlsx";
            return _pathCompatibilityService.BuildUniquePath(tempFolder, baseName, extension);
        }

        internal bool MoveLocalWorkingCaseToFinalPath(string localWorkingPath, string finalCaseWorkbookPath)
        {
            return _pathCompatibilityService.MoveFileSafe(localWorkingPath, finalCaseWorkbookPath);
        }

        internal bool EnsureFolderExists(string folderPath)
        {
            return _pathCompatibilityService.EnsureFolderSafe(folderPath);
        }

        internal string SelectFolderPath(string dialogTitle, string initialDirectory)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Multiselect = false;
                dialog.Title = dialogTitle ?? string.Empty;
                dialog.EnsurePathExists = true;
                dialog.AllowNonFileSystemItems = false;

                string normalizedDirectory = _pathCompatibilityService.NormalizePath(initialDirectory);
                if (!string.IsNullOrWhiteSpace(normalizedDirectory) && Directory.Exists(normalizedDirectory))
                {
                    dialog.InitialDirectory = normalizedDirectory;
                    dialog.DefaultDirectory = normalizedDirectory;
                }

                return dialog.ShowDialog() == CommonFileDialogResult.Ok
                    ? _pathCompatibilityService.NormalizePath(dialog.FileName)
                    : string.Empty;
            }
        }
    }
}
