using System.Linq;
using CaseInfoSystem.ExcelAddIn.UI;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public sealed class KernelNavigationModelsTests
    {
        [Fact]
        public void CreateForSheet_WhenTemplateList_AddsTemplateActionsBeforeReflectTemplate()
        {
            var actions = KernelNavigationDefinitions.CreateForSheet("shMasterList");
            string[] actionIds = actions.Select(action => action.ActionId).ToArray();

            int openTemplateFolderIndex = System.Array.IndexOf(actionIds, "open-template-folder");
            int syncBaseHomeFieldInventoryIndex = System.Array.IndexOf(actionIds, "sync-base-home-field-inventory");
            int reflectTemplateIndex = System.Array.IndexOf(actionIds, "reflect-template");

            Assert.True(openTemplateFolderIndex >= 0);
            Assert.True(syncBaseHomeFieldInventoryIndex >= 0);
            Assert.True(reflectTemplateIndex >= 0);
            Assert.True(openTemplateFolderIndex < syncBaseHomeFieldInventoryIndex);
            Assert.True(openTemplateFolderIndex < reflectTemplateIndex);
            Assert.True(syncBaseHomeFieldInventoryIndex < reflectTemplateIndex);
        }

        [Fact]
        public void CreateForSheet_WhenNotTemplateList_DoesNotAddTemplateActions()
        {
            var actions = KernelNavigationDefinitions.CreateForSheet("shUserData");

            Assert.DoesNotContain(actions, action => action.ActionId == "open-template-folder");
            Assert.DoesNotContain(actions, action => action.ActionId == "sync-base-home-field-inventory");
        }
    }
}
