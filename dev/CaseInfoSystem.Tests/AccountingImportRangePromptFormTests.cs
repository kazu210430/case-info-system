using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class AccountingImportRangePromptFormTests
    {
        [Fact]
        public void Constructor_RemovesLegacyCloseButton_AndKeepsExcelCloseButton()
        {
            using (var form = new AccountingImportRangePromptForm(1, 3))
            {
                Button[] buttons = form.Controls.OfType<Button>().ToArray();

                Assert.DoesNotContain(buttons, button => string.Equals(button.Name, "btnClose", StringComparison.OrdinalIgnoreCase));
                Assert.DoesNotContain(buttons, button => string.Equals(button.Text, "閉じる", StringComparison.Ordinal));
                Assert.Contains(buttons, button => string.Equals(button.Name, "btnExcelClose", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(button.Text, "Excelを閉じる", StringComparison.Ordinal));
                Assert.Contains(buttons, button => string.Equals(button.Name, "btnConfirm", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Type_DoesNotDeclareLegacyUserClosingGuard()
        {
            Type formType = typeof(AccountingImportRangePromptForm);
            const BindingFlags instanceDeclared = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly;
            const BindingFlags instanceNonPublicDeclared = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly;

            Assert.Null(formType.GetField("_allowCloseByButton", instanceNonPublicDeclared));
            Assert.Null(formType.GetMethod("CloseByCode", instanceDeclared));
            Assert.Null(formType.GetMethod("BtnClose_Click", instanceNonPublicDeclared));
            Assert.Null(formType.GetMethod("OnFormClosing", instanceNonPublicDeclared));
            Assert.Null(formType.GetEvent("Canceled", instanceDeclared));
        }

        [Fact]
        public void ImportService_CleansHighlightOnFormClosing_AndDetachesOnFormClosed()
        {
            string source = ReadAppSource("AccountingPaymentHistoryImportService.cs");

            Assert.Contains("form.FormClosing += ActivePromptForm_FormClosing;", source);
            Assert.Contains("form.FormClosed += ActivePromptForm_FormClosed;", source);
            Assert.Contains("CleanupActivePromptHighlightOnce (\"FormClosing\")", source);
            Assert.Contains("CleanupActivePromptHighlightOnce (\"FormClosed\")", source);
            Assert.Contains("_accountingWorkbookService.ClearAccountingImportTargetHighlight (_activePromptWorkbook)", source);
            Assert.Contains("DetachPromptHandlers (form);", source);
            Assert.Contains("ClearActivePromptReferences ();", source);
            Assert.Contains("F15:F20", source);
            Assert.DoesNotContain("form.Canceled +=", source);
            Assert.DoesNotContain("form.CloseByCode", source);
        }

        private static string ReadAppSource(string appFileName)
        {
            string repoRoot = FindRepositoryRoot();
            return File.ReadAllText(Path.Combine(repoRoot, "dev", "CaseInfoSystem.ExcelAddIn", "App", appFileName));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
