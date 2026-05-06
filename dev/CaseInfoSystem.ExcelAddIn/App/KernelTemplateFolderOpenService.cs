using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelTemplateFolderOpenService
    {
        internal sealed class OpenResult
        {
            private OpenResult(bool success, string folderPath, string failureMessage)
            {
                Success = success;
                FolderPath = folderPath ?? string.Empty;
                FailureMessage = failureMessage ?? string.Empty;
            }

            internal bool Success { get; }

            internal string FolderPath { get; }

            internal string FailureMessage { get; }

            internal static OpenResult Succeeded(string folderPath)
            {
                return new OpenResult(true, folderPath, string.Empty);
            }

            internal static OpenResult Failed(string folderPath, string failureMessage)
            {
                return new OpenResult(false, folderPath, failureMessage);
            }
        }

        private const string OpenReason = "KernelCommandService.OpenTemplateFolder";

        private readonly KernelTemplateFolderPathResolver _pathResolver;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly FolderWindowService _folderWindowService;
        private readonly Logger _logger;

        internal KernelTemplateFolderOpenService(
            KernelTemplateFolderPathResolver pathResolver,
            PathCompatibilityService pathCompatibilityService,
            FolderWindowService folderWindowService,
            Logger logger)
        {
            _pathResolver = pathResolver ?? throw new ArgumentNullException(nameof(pathResolver));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _folderWindowService = folderWindowService ?? throw new ArgumentNullException(nameof(folderWindowService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal OpenResult TryOpen(Excel.Workbook kernelWorkbook)
        {
            if (kernelWorkbook == null)
            {
                throw new ArgumentNullException(nameof(kernelWorkbook));
            }

            string folderPath = _pathResolver.ResolveDirectoryForOpen(kernelWorkbook);
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                _logger.Warn("Template folder open failed because the folder path could not be resolved.");
                return OpenResult.Failed(
                    string.Empty,
                    "雛形フォルダを解決できませんでした。Kernel ブックの SYSTEM_ROOT / WORD_TEMPLATE_DIR を確認してください。");
            }

            if (!_pathCompatibilityService.DirectoryExistsSafe(folderPath))
            {
                _logger.Warn("Template folder open failed because the folder does not exist. folder=" + folderPath);
                return OpenResult.Failed(
                    folderPath,
                    "雛形フォルダが見つかりませんでした。"
                    + Environment.NewLine
                    + "フォルダ: "
                    + folderPath);
            }

            if (!_folderWindowService.OpenFolder(folderPath, OpenReason))
            {
                _logger.Warn("Template folder open failed because explorer launch did not succeed. folder=" + folderPath);
                return OpenResult.Failed(
                    folderPath,
                    "雛形フォルダを開けませんでした。ログを確認してください。"
                    + Environment.NewLine
                    + "フォルダ: "
                    + folderPath);
            }

            _logger.Info("Template folder opened. folder=" + folderPath);
            return OpenResult.Succeeded(folderPath);
        }
    }
}
