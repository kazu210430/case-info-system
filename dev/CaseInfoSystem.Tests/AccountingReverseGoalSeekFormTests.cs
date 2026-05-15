using System;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class AccountingReverseGoalSeekFormTests
    {
        [Fact]
        public void Constructor_RemovesLegacyCloseButton_AndKeepsExcelCloseButton()
        {
            using (var form = new AccountingReverseGoalSeekForm())
            {
                Button[] buttons = form.Controls.OfType<Button>().ToArray();

                Assert.DoesNotContain(buttons, button => string.Equals(button.Name, "btnClose", StringComparison.OrdinalIgnoreCase));
                Assert.DoesNotContain(buttons, button => string.Equals(button.Text, "閉じる", StringComparison.Ordinal));
                Assert.Contains(buttons, button => string.Equals(button.Name, "btnExcelClose", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(button.Text, "Excelを閉じる", StringComparison.Ordinal));
                Assert.Contains(buttons, button => string.Equals(button.Name, "btnCalculate", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Type_DoesNotDeclareLegacyUserClosingGuard()
        {
            Type formType = typeof(AccountingReverseGoalSeekForm);
            const BindingFlags instanceNonPublicDeclared = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly;

            Assert.Null(formType.GetField("_allowCloseByButton", instanceNonPublicDeclared));
            Assert.Null(formType.GetMethod("CloseByCode", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly));
            Assert.Null(formType.GetMethod("BtnClose_Click", instanceNonPublicDeclared));
            Assert.Null(formType.GetMethod("OnFormClosing", instanceNonPublicDeclared));
        }
    }
}
