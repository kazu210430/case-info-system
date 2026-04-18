using System;
using System.IO;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class WorkbookFileNameResolver
    {
        internal const string KernelWorkbookBaseName = "\u6848\u4ef6\u60c5\u5831System_Kernel";
        internal const string BaseWorkbookBaseName = "\u6848\u4ef6\u60c5\u5831System_Base";
        internal const string CaseWorkbookBaseName = "\u6848\u4ef6\u60c5\u5831";

        private static readonly string MacroEnabledWorkbookExtension = new string(new[] { '.', 'x', 'l', 's', 'm' });
        private static readonly string OpenXmlWorkbookExtension = new string(new[] { '.', 'x', 'l', 's', 'x' });
        private static readonly string[] SupportedMainWorkbookExtensions =
        {
            MacroEnabledWorkbookExtension,
            OpenXmlWorkbookExtension
        };

        internal static bool IsKernelWorkbookName(string workbookName)
        {
            return IsWorkbookName(workbookName, KernelWorkbookBaseName);
        }

        internal static bool IsBaseWorkbookName(string workbookName)
        {
            return IsWorkbookName(workbookName, BaseWorkbookBaseName);
        }

        internal static bool IsSupportedCaseExtension(string extension)
        {
            return HasSupportedExtension(extension);
        }

        internal static string[] GetSupportedMainWorkbookExtensions()
        {
            var extensions = new string[SupportedMainWorkbookExtensions.Length];
            Array.Copy(SupportedMainWorkbookExtensions, extensions, SupportedMainWorkbookExtensions.Length);
            return extensions;
        }

        internal static string BuildKernelWorkbookName(string extension)
        {
            return KernelWorkbookBaseName + NormalizeMainWorkbookExtension(extension);
        }

        internal static string BuildBaseWorkbookName(string extension)
        {
            return BaseWorkbookBaseName + NormalizeMainWorkbookExtension(extension);
        }

        internal static string BuildCaseWorkbookName(string customerName, string extension)
        {
            string normalizedCustomer = App.KernelNamingService.SanitizeFileNameForCase(customerName);
            string resolvedExtension = NormalizeMainWorkbookExtension(extension);
            return string.IsNullOrWhiteSpace(normalizedCustomer)
                ? CaseWorkbookBaseName + resolvedExtension
                : CaseWorkbookBaseName + "_" + normalizedCustomer + resolvedExtension;
        }

        internal static string ResolveExistingKernelWorkbookPath(string systemRoot, PathCompatibilityService pathCompatibilityService)
        {
            return ResolveExistingWorkbookPath(systemRoot, KernelWorkbookBaseName, pathCompatibilityService);
        }

        internal static string ResolveExistingBaseWorkbookPath(string systemRoot, PathCompatibilityService pathCompatibilityService)
        {
            return ResolveExistingWorkbookPath(systemRoot, BaseWorkbookBaseName, pathCompatibilityService);
        }

        internal static string GetWorkbookExtensionOrDefault(string workbookPath)
        {
            string extension = Path.GetExtension(workbookPath) ?? string.Empty;
            return NormalizeMainWorkbookExtension(extension);
        }

        private static string ResolveExistingWorkbookPath(string systemRoot, string workbookBaseName, PathCompatibilityService pathCompatibilityService)
        {
            if (pathCompatibilityService == null)
            {
                throw new ArgumentNullException(nameof(pathCompatibilityService));
            }

            string normalizedRoot = pathCompatibilityService.NormalizePath(systemRoot);
            if (string.IsNullOrWhiteSpace(normalizedRoot))
            {
                return string.Empty;
            }

            for (int index = 0; index < SupportedMainWorkbookExtensions.Length; index++)
            {
                string candidate = pathCompatibilityService.CombinePath(normalizedRoot, workbookBaseName + SupportedMainWorkbookExtensions[index]);
                if (pathCompatibilityService.FileExistsSafe(candidate))
                {
                    return pathCompatibilityService.NormalizePath(candidate);
                }
            }

            return string.Empty;
        }

        private static bool IsWorkbookName(string workbookName, string expectedBaseName)
        {
            string normalizedName = (workbookName ?? string.Empty).Trim();
            if (normalizedName.Length == 0)
            {
                return false;
            }

            string actualBaseName = Path.GetFileNameWithoutExtension(normalizedName) ?? string.Empty;
            if (!string.Equals(actualBaseName, expectedBaseName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return HasSupportedExtension(Path.GetExtension(normalizedName));
        }

        private static bool HasSupportedExtension(string extension)
        {
            string normalizedExtension = (extension ?? string.Empty).Trim();
            for (int index = 0; index < SupportedMainWorkbookExtensions.Length; index++)
            {
                if (string.Equals(normalizedExtension, SupportedMainWorkbookExtensions[index], StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeMainWorkbookExtension(string extension)
        {
            string normalizedExtension = (extension ?? string.Empty).Trim();
            return string.Equals(normalizedExtension, OpenXmlWorkbookExtension, StringComparison.OrdinalIgnoreCase)
                ? OpenXmlWorkbookExtension
                : MacroEnabledWorkbookExtension;
        }
    }
}
