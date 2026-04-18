using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class WorkbookRoleResolver : IWorkbookRoleResolver
    {
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string RolePropertyName = "ROLE";
        private const string CaseRoleName = "CASE";
        private const string BaseRoleName = "BASE";
        private const string SourceCasePathPropertyName = "SOURCE_CASE_PATH";
        private const string SourceKernelPathPropertyName = "SOURCE_KERNEL_PATH";

        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly HashSet<string> _knownCaseWorkbookIdentities;

        /// <summary>
        internal WorkbookRoleResolver(ExcelInteropService excelInteropService, PathCompatibilityService pathCompatibilityService)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _knownCaseWorkbookIdentities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        public WorkbookRole Resolve(Excel.Workbook workbook)
        {
            if (IsKernelWorkbook(workbook))
            {
                return WorkbookRole.Kernel;
            }

            if (IsAccountingWorkbook(workbook))
            {
                return WorkbookRole.Accounting;
            }

            if (IsCaseWorkbook(workbook))
            {
                return WorkbookRole.Case;
            }

            return WorkbookRole.Unknown;
        }

        /// <summary>
        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            string workbookName = _excelInteropService.GetWorkbookName(workbook);
            return WorkbookFileNameResolver.IsKernelWorkbookName(workbookName);
        }

        /// <summary>
        internal bool IsBaseWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            string workbookName = _excelInteropService.GetWorkbookName(workbook);
            if (WorkbookFileNameResolver.IsBaseWorkbookName(workbookName))
            {
                return true;
            }

            string role = _excelInteropService.TryGetDocumentProperty(workbook, RolePropertyName);
            return string.Equals(role, BaseRoleName, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        internal bool IsCaseWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null || IsKernelWorkbook(workbook) || IsBaseWorkbook(workbook))
            {
                return false;
            }

            if (HasAccountingSheets(workbook))
            {
                return false;
            }

            if (IsKnownCaseWorkbook(workbook))
            {
                return true;
            }

            string role = _excelInteropService.TryGetDocumentProperty(workbook, RolePropertyName);
            if (!string.Equals(role, CaseRoleName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string systemRoot = _excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName);
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                return false;
            }

            string workbookName = _excelInteropService.GetWorkbookName(workbook);
            string extension = Path.GetExtension(workbookName) ?? string.Empty;
            return WorkbookFileNameResolver.IsSupportedCaseExtension(extension);
        }

        internal void RegisterKnownCaseWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            RegisterKnownCaseIdentity(_excelInteropService.GetWorkbookFullName(workbook));
            RegisterKnownCaseIdentity(_excelInteropService.GetWorkbookPath(workbook));
            RegisterKnownCaseIdentity(_excelInteropService.GetWorkbookName(workbook));
        }

        internal void RegisterKnownCasePath(string path)
        {
            RegisterKnownCaseIdentity(path);
            RegisterKnownCaseIdentity(_pathCompatibilityService.GetFileNameFromPath(path));
        }

        internal void RemoveKnownWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            RemoveKnownCaseIdentity(_excelInteropService.GetWorkbookFullName(workbook));
            RemoveKnownCaseIdentity(_excelInteropService.GetWorkbookPath(workbook));
            RemoveKnownCaseIdentity(_excelInteropService.GetWorkbookName(workbook));
        }

        /// <summary>
        internal bool IsAccountingWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null || IsKernelWorkbook(workbook) || IsBaseWorkbook(workbook))
            {
                return false;
            }

            string workbookName = _excelInteropService.GetWorkbookName(workbook);
            string extension = Path.GetExtension(workbookName) ?? string.Empty;
            if (!WorkbookFileNameResolver.IsSupportedCaseExtension(extension))
            {
                return false;
            }

            string workbookKind = (_excelInteropService.TryGetDocumentProperty(workbook, AccountingSetSpec.WorkbookKindPropertyName) ?? string.Empty).Trim();
            if (string.Equals(workbookKind, AccountingSetSpec.WorkbookKindAccountingSetValue, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            bool hasSourceProperty =
                !string.IsNullOrWhiteSpace(_excelInteropService.TryGetDocumentProperty(workbook, SourceCasePathPropertyName))
                || !string.IsNullOrWhiteSpace(_excelInteropService.TryGetDocumentProperty(workbook, SourceKernelPathPropertyName));

            return HasAccountingSheets(workbook) && (hasSourceProperty || !IsKnownCaseWorkbook(workbook));
        }

        /// <summary>
        public string ResolveSystemRoot(Excel.Workbook workbook)
        {
            return _excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName);
        }

        private bool IsKnownCaseWorkbook(Excel.Workbook workbook)
        {
            return ContainsKnownCaseIdentity(_excelInteropService.GetWorkbookFullName(workbook))
                || ContainsKnownCaseIdentity(_excelInteropService.GetWorkbookPath(workbook))
                || ContainsKnownCaseIdentity(_excelInteropService.GetWorkbookName(workbook));
        }

        private bool ContainsKnownCaseIdentity(string identity)
        {
            string normalized = NormalizeIdentity(identity);
            return normalized.Length > 0 && _knownCaseWorkbookIdentities.Contains(normalized);
        }

        private void RegisterKnownCaseIdentity(string identity)
        {
            string normalized = NormalizeIdentity(identity);
            if (normalized.Length > 0)
            {
                _knownCaseWorkbookIdentities.Add(normalized);
            }
        }

        private void RemoveKnownCaseIdentity(string identity)
        {
            string normalized = NormalizeIdentity(identity);
            if (normalized.Length > 0)
            {
                _knownCaseWorkbookIdentities.Remove(normalized);
            }
        }

        private string NormalizeIdentity(string identity)
        {
            string normalized = _pathCompatibilityService.NormalizePath(identity);
            if (normalized.Length > 0)
            {
                return normalized;
            }

            return (identity ?? string.Empty).Trim();
        }

        private static bool HasWorksheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            try
            {
                return workbook.Worksheets[sheetName] is Excel.Worksheet;
            }
            catch
            {
                return false;
            }
        }

        private static bool HasAccountingSheets(Excel.Workbook workbook)
        {
            return HasWorksheet(workbook, Domain.AccountingSetSpec.InvoiceSheetName)
                && HasWorksheet(workbook, Domain.AccountingSetSpec.PaymentHistorySheetName)
                && HasWorksheet(workbook, Domain.AccountingSetSpec.AccountingRequestSheetName);
        }
    }
}
