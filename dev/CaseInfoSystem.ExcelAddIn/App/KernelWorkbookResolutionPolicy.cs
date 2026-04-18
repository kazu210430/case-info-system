using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class KernelWorkbookResolutionPolicy
    {
        internal static string ResolveKernelWorkbookPath(
            bool hasOpenKernelWorkbook,
            string systemRoot,
            Func<string, string> resolvePath)
        {
            if (hasOpenKernelWorkbook)
            {
                return string.Empty;
            }

            if (string.IsNullOrWhiteSpace(systemRoot) || resolvePath == null)
            {
                return string.Empty;
            }

            return resolvePath(systemRoot) ?? string.Empty;
        }
    }
}
