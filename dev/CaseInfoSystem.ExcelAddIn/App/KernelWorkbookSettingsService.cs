using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.WindowsAPICodePack.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookSettingsService
    {
        private const string DefaultRootPropertyName = "DEFAULT_ROOT";
        private const string NameRuleAPropertyName = "NAME_RULE_A";
        private const string NameRuleBPropertyName = "NAME_RULE_B";
        private const string DefaultNameRuleA = "YYYY";
        private const string DefaultNameRuleB = "DOC";
        private const string DefaultRootDialogTitle = "\u65B0\u898F\u30D5\u30A9\u30EB\u30C0\u306E\u89AA\uFF08\u4FDD\u5B58\u5148\uFF09\u30D5\u30A9\u30EB\u30C0\u3092\u9078\u629E\u3057\u3066\u304F\u3060\u3055\u3044";

        internal string LoadNameRuleA(Excel.Workbook workbook)
        {
            string nameRuleA = TryGetDocumentProperty(workbook, NameRuleAPropertyName);
            return KernelNamingService.NormalizeNameRuleA(string.IsNullOrWhiteSpace(nameRuleA) ? DefaultNameRuleA : nameRuleA);
        }

        internal string LoadNameRuleB(Excel.Workbook workbook)
        {
            string nameRuleB = TryGetDocumentProperty(workbook, NameRuleBPropertyName);
            return KernelNamingService.NormalizeNameRuleB(string.IsNullOrWhiteSpace(nameRuleB) ? DefaultNameRuleB : nameRuleB);
        }

        internal string LoadDefaultRoot(Excel.Workbook workbook)
        {
            return TryGetDocumentProperty(workbook, DefaultRootPropertyName);
        }

        internal void SaveNameRuleA(Excel.Workbook workbook, string ruleA)
        {
            SetDocumentProperty(workbook, NameRuleAPropertyName, ruleA);
        }

        internal void SaveNameRuleB(Excel.Workbook workbook, string ruleB)
        {
            SetDocumentProperty(workbook, NameRuleBPropertyName, ruleB);
        }

        internal string SelectDefaultRoot(Excel.Workbook workbook)
        {
            string selectedPath = SelectFolderPath(DefaultRootDialogTitle, LoadDefaultRoot(workbook));
            if (string.IsNullOrWhiteSpace(selectedPath))
            {
                return null;
            }

            SetDocumentProperty(workbook, DefaultRootPropertyName, selectedPath);
            return selectedPath;
        }

        private static string SelectFolderPath(string dialogTitle, string initialDirectory)
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Multiselect = false;
                dialog.Title = dialogTitle;
                dialog.EnsurePathExists = true;
                dialog.AllowNonFileSystemItems = false;

                if (!string.IsNullOrWhiteSpace(initialDirectory) && Directory.Exists(initialDirectory))
                {
                    dialog.InitialDirectory = initialDirectory;
                    dialog.DefaultDirectory = initialDirectory;
                }

                if (dialog.ShowDialog() != CommonFileDialogResult.Ok)
                {
                    return null;
                }

                return dialog.FileName;
            }
        }

        private static string TryGetDocumentProperty(Excel.Workbook workbook, string propertyName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            object properties = null;
            object property = null;
            try
            {
                properties = workbook.CustomDocumentProperties;
                dynamic dynamicProperties = properties;
                property = dynamicProperties[propertyName];
                dynamic dynamicProperty = property;
                object value = dynamicProperty.Value;
                return Convert.ToString(value) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(property);
                ReleaseComObject(properties);
            }
        }

        private static void SetDocumentProperty(Excel.Workbook workbook, string propertyName, string value)
        {
            object properties = workbook.CustomDocumentProperties;
            try
            {
                dynamic dynamicProperties = properties;
                try
                {
                    dynamicProperties[propertyName].Value = value;
                }
                catch
                {
                    const int MsoPropertyTypeString = 4;
                    dynamicProperties.Add(propertyName, false, MsoPropertyTypeString, value);
                }
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch
                {
                }
            }
        }
    }
}
