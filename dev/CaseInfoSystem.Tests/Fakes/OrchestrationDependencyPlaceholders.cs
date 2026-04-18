using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class ExcelInteropService
    {
        internal Excel.Workbook GetActiveWorkbook() => null;

        internal Excel.Window GetActiveWindow() => null;

        internal string GetWorkbookFullName(Excel.Workbook workbook) => workbook == null ? string.Empty : workbook.FullName ?? string.Empty;

        internal string GetWorkbookName(Excel.Workbook workbook) => workbook == null ? string.Empty : workbook.Name ?? string.Empty;

        internal string GetWorkbookPath(Excel.Workbook workbook) => workbook == null ? string.Empty : workbook.Path ?? string.Empty;

        internal string TryGetDocumentProperty(Excel.Workbook workbook, string propertyName)
        {
            if (workbook?.CustomDocumentProperties is IDictionary<string, string> properties
                && properties.TryGetValue(propertyName ?? string.Empty, out string value))
            {
                return value ?? string.Empty;
            }

            return string.Empty;
        }

        internal void SetDocumentProperty(Excel.Workbook workbook, string propertyName, string value)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return;
            }

            if (!(workbook.CustomDocumentProperties is IDictionary<string, string> properties))
            {
                properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                workbook.CustomDocumentProperties = properties;
            }

            properties[propertyName] = value ?? string.Empty;
        }

        internal Excel.Window GetFirstVisibleWindow(Excel.Workbook workbook) => workbook == null ? null : workbook.Windows.FirstOrDefault(window => window.Visible);

        internal string GetActiveSheetCodeName(Excel.Workbook workbook) => workbook?.ActiveSheet?.CodeName ?? string.Empty;

        internal Excel.Worksheet FindWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName)
        {
            return workbook?.Worksheets.FirstOrDefault(worksheet => string.Equals(worksheet?.CodeName, sheetCodeName, StringComparison.OrdinalIgnoreCase));
        }

        internal Excel.Workbook FindOpenWorkbook(string workbookPath) => null;

        internal bool ActivateWorkbook(Excel.Workbook workbook) => true;

        internal bool ActivateWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName) => true;
    }

    internal sealed class PathCompatibilityService
    {
        internal string NormalizePath(string path) => (path ?? string.Empty).Trim();

        internal bool FileExistsSafe(string path) => !string.IsNullOrWhiteSpace(path);

        internal bool DirectoryExistsSafe(string path) => !string.IsNullOrWhiteSpace(path);

        internal string GetFileNameFromPath(string path) => Path.GetFileName(path ?? string.Empty) ?? string.Empty;

        internal string CombinePath(string left, string right) => Path.Combine(left ?? string.Empty, right ?? string.Empty);
    }

    internal sealed class WorkbookRoleResolver
    {
        internal bool IsBaseWorkbook(Excel.Workbook workbook) => false;

        internal bool IsCaseWorkbook(Excel.Workbook workbook) => false;

        internal void RegisterKnownCaseWorkbook(Excel.Workbook workbook)
        {
        }

        internal void RemoveKnownWorkbook(Excel.Workbook workbook)
        {
        }
    }

    internal sealed class ExcelWindowRecoveryService
    {
        internal void EnsureApplicationVisible(string reason, string workbookFullName)
        {
        }

        internal bool TryRestoreMainWindow(bool bringToFront) => true;

        internal bool TryRestoreWorkbookWindow(Excel.Workbook workbook, bool bringToFront) => true;

        internal bool TryRecoverWorkbookWindow(Excel.Workbook workbook, string reason, bool bringToFront) => true;
    }

    internal sealed class UserErrorService
    {
        internal void ShowUserError(string context, Exception ex)
        {
        }
    }

    internal sealed class TaskPaneSnapshotBuilderService
    {
        internal sealed class TaskPaneBuildResult
        {
            internal TaskPaneBuildResult(string snapshotText, bool updatedCaseSnapshotCache)
            {
                SnapshotText = snapshotText ?? string.Empty;
                UpdatedCaseSnapshotCache = updatedCaseSnapshotCache;
            }

            internal string SnapshotText { get; }

            internal bool UpdatedCaseSnapshotCache { get; }
        }

        internal Func<Excel.Workbook, TaskPaneBuildResult> OnBuildSnapshotText { get; set; }

        internal TaskPaneBuildResult BuildSnapshotText(Excel.Workbook workbook)
        {
            return OnBuildSnapshotText != null
                ? OnBuildSnapshotText(workbook)
                : new TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: false);
        }
    }
}

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class DocumentCommandService
    {
        internal void Execute(Excel.Workbook workbook, string actionKind, string key)
        {
        }
    }

    internal sealed class DocumentEligibilityDiagnosticsService
    {
    }

    internal sealed class DocumentMasterCatalogDiagnosticsService
    {
    }

    internal sealed class DocumentNamePromptService
    {
        internal bool TryPrepare(Excel.Workbook workbook, string key, out DocumentNameOverrideScope scope)
        {
            scope = null;
            return true;
        }
    }

    internal sealed class KernelCommandService
    {
        internal void Execute(WorkbookContext context, string actionId)
        {
        }
    }

    internal sealed class AccountingSheetCommandService
    {
        internal void Execute(WorkbookContext context, string actionId)
        {
        }

        internal void ShowInstallmentSchedule(Excel.Workbook workbook)
        {
        }

        internal void ShowPaymentHistory(Excel.Workbook workbook)
        {
        }

        internal void RunReverseGoalSeek(Excel.Workbook workbook)
        {
        }
    }

    internal sealed class AccountingInternalCommandService
    {
        internal void ExecuteImportPaymentHistory(Excel.Workbook workbook)
        {
        }

        internal void Execute(WorkbookContext context, string actionId)
        {
        }
    }

    internal sealed class TransientPaneSuppressionService
    {
        internal bool IsSuppressed(Excel.Workbook workbook) => false;
    }

    internal sealed class KernelWorkbookLifecycleService
    {
        internal bool RequestManagedCloseFromHomeExit(Excel.Workbook workbook) => true;
    }
}
