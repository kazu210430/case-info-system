using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class CaseClosePromptService
    {
        private const string RolePropertyName = "ROLE";
        private const string CaseRoleName = "CASE";
        private const string BaseRoleName = "BASE";
        private const string CreatedCaseFolderOfferPromptTitle = "案件情報System (CASE)";
        private const string CreatedCaseFolderOfferPromptMessage = "保存しました\r\n保存先フォルダを開きますか？";
        private const string CreatedCaseFolderOfferOpenReason = "CaseWorkbookLifecycleService.PostCloseCreatedCaseFolderOffer";

        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly FolderWindowService _folderWindowService;
        private readonly Logger _logger;

        internal CaseClosePromptService(
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            FolderWindowService folderWindowService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _folderWindowService = folderWindowService ?? throw new ArgumentNullException(nameof(folderWindowService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal DialogResult ShowClosePrompt(Excel.Workbook workbook)
        {
            return MessageBox.Show(
                "保存しますか？",
                BuildCloseDialogTitle(workbook),
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);
        }

        internal string GetCloseDialogTitle(Excel.Workbook workbook)
        {
            return BuildCloseDialogTitle(workbook);
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

        internal void TryPromptToOpenCreatedCaseFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return;
            }

            if (!DirectoryExistsSafe(folderPath))
            {
                _logger.Info("Created CASE folder offer prompt skipped because folder does not exist. folderPath=" + folderPath);
                return;
            }

            try
            {
                DialogResult answer = MessageBox.Show(
                    CreatedCaseFolderOfferPromptMessage,
                    CreatedCaseFolderOfferPromptTitle,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                if (answer != DialogResult.Yes)
                {
                    _logger.Info("Created CASE folder offer prompt dismissed without opening folder. folderPath=" + folderPath + ", answer=" + answer.ToString());
                    return;
                }

                _folderWindowService.OpenFolder(folderPath, CreatedCaseFolderOfferOpenReason);
            }
            catch (Exception ex)
            {
                _logger.Error("Created CASE folder offer prompt failed.", ex);
            }
        }

        private string BuildCloseDialogTitle(Excel.Workbook workbook)
        {
            string role = (_excelInteropService.TryGetDocumentProperty(workbook, RolePropertyName) ?? string.Empty)
                .Trim()
                .ToUpperInvariant();
            if (role == CaseRoleName)
            {
                return "案件情報System (CASE)";
            }

            if (role == BaseRoleName)
            {
                return "案件情報System (BASE)";
            }

            return "案件情報System";
        }
    }
}
