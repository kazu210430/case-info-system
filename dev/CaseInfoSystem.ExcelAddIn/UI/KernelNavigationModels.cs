using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    internal sealed class KernelNavigationActionDefinition
    {
        internal string ActionId { get; }

        internal string Caption { get; }

        internal string SectionTitle { get; }

        internal bool IsEnabled { get; }

        internal bool IsCurrentDisplay { get; }

        internal KernelNavigationActionDefinition(string actionId, string caption, string sectionTitle, bool isEnabled, bool isCurrentDisplay)
        {
            ActionId = actionId ?? string.Empty;
            Caption = caption ?? string.Empty;
            SectionTitle = sectionTitle ?? string.Empty;
            IsEnabled = isEnabled;
            IsCurrentDisplay = isCurrentDisplay;
        }
    }

    internal static class KernelNavigationActionIds
    {
        internal const string OpenHome = "open-home";
        internal const string OpenUserInfo = "open-user-info";
        internal const string OpenTemplateList = "open-template-list";
        internal const string OpenCaseList = "open-case-list";
        internal const string RegisterUserInfo = "register-user-info";
        internal const string OpenTemplateFolder = "open-template-folder";
        internal const string ReflectTemplate = "reflect-template";
        internal const string SyncBaseHomeFieldInventory = "sync-base-home-field-inventory";
    }

    internal static class KernelNavigationDefinitions
    {
        private const string SectionScreen = "画面切替";
        private const string SectionAction = "実行";

        internal static IReadOnlyList<KernelNavigationActionDefinition> CreateForSheet(string activeSheetCodeName)
        {
            bool isUserData = string.Equals(activeSheetCodeName, "shUserData", StringComparison.OrdinalIgnoreCase);
            bool isTemplateList = string.Equals(activeSheetCodeName, "shMasterList", StringComparison.OrdinalIgnoreCase);
            bool isCaseList = string.Equals(activeSheetCodeName, "shCaseList", StringComparison.OrdinalIgnoreCase);
            List<KernelNavigationActionDefinition> definitions = new List<KernelNavigationActionDefinition>();

            if (isUserData)
            {
                definitions.Add(new KernelNavigationActionDefinition("register-user-info", "ユーザー情報登録", SectionAction, true, false));
            }

            if (isTemplateList)
            {
                definitions.Add(new KernelNavigationActionDefinition("open-template-folder", "雛形フォルダを開く", SectionAction, true, false));
                definitions.Add(new KernelNavigationActionDefinition("reflect-template", "雛形登録・更新", SectionAction, true, false));
            }

            definitions.Add(new KernelNavigationActionDefinition("open-home", "案件情報 作成 HOME", SectionScreen, true, false));
            definitions.Add(new KernelNavigationActionDefinition("open-user-info", BuildScreenCaption("ユーザー情報", isUserData), SectionScreen, !isUserData, isUserData));
            definitions.Add(new KernelNavigationActionDefinition("open-template-list", BuildScreenCaption("雛形一覧", isTemplateList), SectionScreen, !isTemplateList, isTemplateList));
            definitions.Add(new KernelNavigationActionDefinition("open-case-list", BuildScreenCaption("案件一覧", isCaseList), SectionScreen, !isCaseList, isCaseList));
            return definitions;
        }

        private static string BuildScreenCaption(string baseCaption, bool isCurrentDisplay)
        {
            return isCurrentDisplay ? (baseCaption ?? string.Empty) + "（表示中）" : (baseCaption ?? string.Empty);
        }
    }

    internal sealed class KernelNavigationActionEventArgs : EventArgs
    {
        internal string ActionId { get; }

        internal KernelNavigationActionEventArgs(string actionId)
        {
            ActionId = actionId ?? string.Empty;
        }
    }
}
