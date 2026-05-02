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

        private readonly ExcelInteropService _excelInteropService;

        internal CaseClosePromptService(
            ExcelInteropService excelInteropService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
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

        internal DialogResult ShowCreatedCaseFolderOfferPrompt()
        {
            return MessageBox.Show(
                CreatedCaseFolderOfferPromptMessage,
                CreatedCaseFolderOfferPromptTitle,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);
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
