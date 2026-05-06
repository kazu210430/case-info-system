using System.Linq;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class KernelNavigationModelsTests
    {
        [Fact]
        public void CreateForSheet_WhenTemplateList_AddsOpenTemplateFolderBeforeReflectTemplate()
        {
            var actions = KernelNavigationDefinitions.CreateForSheet("shMasterList");
            string[] actionIds = actions.Select(action => action.ActionId).ToArray();

            int openTemplateFolderIndex = System.Array.IndexOf(actionIds, "open-template-folder");
            int reflectTemplateIndex = System.Array.IndexOf(actionIds, "reflect-template");

            Assert.True(openTemplateFolderIndex >= 0);
            Assert.True(reflectTemplateIndex >= 0);
            Assert.True(openTemplateFolderIndex < reflectTemplateIndex);
        }

        [Fact]
        public void CreateForSheet_WhenNotTemplateList_DoesNotAddOpenTemplateFolder()
        {
            var actions = KernelNavigationDefinitions.CreateForSheet("shUserData");

            Assert.DoesNotContain(actions, action => action.ActionId == "open-template-folder");
        }
    }
}
