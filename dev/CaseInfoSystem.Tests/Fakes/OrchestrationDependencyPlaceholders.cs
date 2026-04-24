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
        private readonly Excel.Application _application;

        internal ExcelInteropService()
        {
        }

        internal ExcelInteropService(Excel.Application application, Logger logger, PathCompatibilityService pathCompatibilityService)
        {
            _application = application;
        }

        internal Func<Excel.Workbook, Excel.Window> OnGetActiveWindow { get; set; }

        internal Func<CaseInfoSystem.ExcelAddIn.Domain.CaseContext, bool> OnTryNormalizeCaseListRowHeight { get; set; }

        internal Func<IEnumerable<Excel.Workbook>> OnGetOpenWorkbooks { get; set; }

        internal Excel.Workbook GetActiveWorkbook() => _application?.ActiveWorkbook;

        internal Excel.Window GetActiveWindow()
        {
            if (OnGetActiveWindow != null)
            {
                return OnGetActiveWindow(null);
            }

            return _application?.ActiveWindow;
        }

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

        internal bool TryNormalizeCaseListRowHeight(CaseInfoSystem.ExcelAddIn.Domain.CaseContext context)
        {
            return OnTryNormalizeCaseListRowHeight == null || OnTryNormalizeCaseListRowHeight(context);
        }

        internal Excel.Workbook FindOpenWorkbook(string workbookPath) => null;

        internal IEnumerable<Excel.Workbook> GetOpenWorkbooks()
        {
            return OnGetOpenWorkbooks != null
                ? OnGetOpenWorkbooks()
                : Enumerable.Empty<Excel.Workbook>();
        }

        internal bool ActivateWorkbook(Excel.Workbook workbook) => true;

        internal bool ActivateWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName) => true;
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
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;

        internal ExcelWindowRecoveryService()
        {
        }

        internal ExcelWindowRecoveryService(Excel.Application application, ExcelInteropService excelInteropService, Logger logger)
        {
            _application = application;
            _excelInteropService = excelInteropService;
        }

        internal void EnsureApplicationVisible(string reason, string workbookFullName)
        {
            if (_application != null)
            {
                _application.Visible = true;
            }
        }

        internal void LogWorkbookWindowSnapshot(Excel.Workbook workbook, string reason, string stage)
        {
        }

        internal bool NormalizeWorkbookWindows(Excel.Workbook workbook, string reason, bool ensurePrimaryVisible, bool activatePrimary, bool bringToFront)
        {
            if (workbook == null)
            {
                return false;
            }

            Excel.Window primaryWindow = null;
            if (_excelInteropService != null
                && string.Equals(_excelInteropService.GetWorkbookFullName(_excelInteropService.GetActiveWorkbook()), workbook.FullName, StringComparison.OrdinalIgnoreCase))
            {
                primaryWindow = _excelInteropService.GetActiveWindow();
            }

            if (primaryWindow == null)
            {
                primaryWindow = _excelInteropService == null
                    ? null
                    : _excelInteropService.GetFirstVisibleWindow(workbook);
            }

            if (primaryWindow == null && workbook.Windows.Count > 0)
            {
                primaryWindow = workbook.Windows[1];
            }

            for (int index = workbook.Windows.Count; index >= 1; index--)
            {
                Excel.Window window = workbook.Windows[index];
                if (window == null || ReferenceEquals(window, primaryWindow))
                {
                    continue;
                }

                window.Close();
            }

            if (primaryWindow == null)
            {
                return false;
            }

            if (ensurePrimaryVisible)
            {
                primaryWindow.Visible = true;
            }

            if (activatePrimary)
            {
                primaryWindow.Activate();
            }

            return true;
        }

        internal bool TryRestoreMainWindow(bool bringToFront) => true;

        internal bool TryRestoreWorkbookWindow(Excel.Workbook workbook, bool bringToFront) => true;

        internal bool TryRecoverWorkbookWindow(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowCore(workbook, reason, allowWindowCreation: true);
        }

        internal bool TryRecoverWorkbookWindowUsingExistingWindows(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowCore(workbook, reason, allowWindowCreation: false);
        }

        internal bool TryRecoverWorkbookWindowWithoutShowing(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowCore(workbook, reason, allowWindowCreation: true);
        }

        internal bool TryRecoverWorkbookWindowWithoutShowingUsingExistingWindows(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            return TryRecoverWorkbookWindowCore(workbook, reason, allowWindowCreation: false);
        }

        private bool TryRecoverWorkbookWindowCore(Excel.Workbook workbook, string reason, bool allowWindowCreation)
        {
            if (workbook == null)
            {
                return false;
            }

            EnsureApplicationVisible(reason, workbook.FullName);

            Excel.Window window = _excelInteropService == null
                ? null
                : _excelInteropService.GetFirstVisibleWindow(workbook);
            if (window == null
                && _excelInteropService != null
                && string.Equals(_excelInteropService.GetWorkbookFullName(_excelInteropService.GetActiveWorkbook()), workbook.FullName, StringComparison.OrdinalIgnoreCase))
            {
                window = _excelInteropService.GetActiveWindow();
            }
            if (window == null)
            {
                workbook.Activate();
                if (workbook.Windows.Count > 0)
                {
                    window = workbook.Windows[1];
                }
                else if (_excelInteropService != null
                    && string.Equals(_excelInteropService.GetWorkbookFullName(_excelInteropService.GetActiveWorkbook()), workbook.FullName, StringComparison.OrdinalIgnoreCase))
                {
                    window = _excelInteropService.GetActiveWindow();
                }
            }

            if (window == null)
            {
                return false;
            }

            window.Visible = true;
            window.Activate();
            return true;
        }
    }

    internal sealed class UserErrorService
    {
        internal void ShowUserError(string context, Exception ex)
        {
        }
    }

    internal sealed class FolderWindowService
    {
        internal Action<string, string> OnOpenFolder { get; set; }

        internal Func<string, string, IntPtr> OnOpenFolderAndWait { get; set; }

        internal void OpenFolder(string folderPath, string reason)
        {
            OnOpenFolder?.Invoke(folderPath, reason);
        }

        internal IntPtr OpenFolderAndWait(string folderPath, string reason)
        {
            return OnOpenFolderAndWait == null
                ? IntPtr.Zero
                : OnOpenFolderAndWait(folderPath, reason);
        }
    }

    internal sealed class CaseWorkbookOpenStrategy
    {
        internal Func<string, HiddenCaseWorkbookSession> OnOpenHiddenWorkbook { get; set; }

        internal HiddenCaseWorkbookSession OpenHiddenWorkbook(string caseWorkbookPath)
        {
            return OnOpenHiddenWorkbook == null
                ? new HiddenCaseWorkbookSession(new Excel.Application(), new Excel.Workbook { FullName = caseWorkbookPath ?? string.Empty }, "legacy-isolated")
                : OnOpenHiddenWorkbook(caseWorkbookPath);
        }

        internal sealed class HiddenCaseWorkbookSession
        {
            private readonly Action _closeAction;
            private bool _closed;

            internal HiddenCaseWorkbookSession(Excel.Application application, Excel.Workbook workbook, string routeName = "legacy-isolated", Action closeAction = null)
            {
                Application = application;
                Workbook = workbook;
                RouteName = routeName ?? string.Empty;
                _closeAction = closeAction ?? (() =>
                {
                    Workbook.Close(false, null, null);
                    Application.Quit();
                });
            }

            internal Excel.Application Application { get; }

            internal Excel.Workbook Workbook { get; }

            internal string RouteName { get; }

            internal void Close()
            {
                ExecuteClose();
            }

            internal void Abort()
            {
                ExecuteClose();
            }

            private void ExecuteClose()
            {
                if (_closed)
                {
                    return;
                }

                _closeAction();
                _closed = true;
            }
        }
    }

    internal sealed class CaseWorkbookInitializer
    {
        internal Action<Excel.Workbook, Excel.Workbook, KernelCaseCreationPlan> OnInitializeForVisibleCreate { get; set; }

        internal Action<Excel.Workbook, Excel.Workbook, KernelCaseCreationPlan> OnInitializeForHiddenCreate { get; set; }

        internal void InitializeForVisibleCreate(Excel.Workbook kernelWorkbook, Excel.Workbook caseWorkbook, KernelCaseCreationPlan plan)
        {
            if (OnInitializeForVisibleCreate != null)
            {
                OnInitializeForVisibleCreate(kernelWorkbook, caseWorkbook, plan);
                return;
            }

            OnInitializeForHiddenCreate?.Invoke(kernelWorkbook, caseWorkbook, plan);
        }

        internal void InitializeForHiddenCreate(Excel.Workbook kernelWorkbook, Excel.Workbook caseWorkbook, KernelCaseCreationPlan plan)
        {
            OnInitializeForHiddenCreate?.Invoke(kernelWorkbook, caseWorkbook, plan);
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
    internal sealed class KernelCasePathService
    {
        internal Func<Excel.Workbook, string> OnResolveSystemRoot { get; set; }

        internal Func<string, string> OnResolveBaseWorkbookPath { get; set; }

        internal Func<KernelCaseCreationRequest, string, string> OnResolveCaseFolderPath { get; set; }

        internal Func<string, bool> OnEnsureFolderExists { get; set; }

        internal Func<string, string> OnResolveCaseWorkbookExtension { get; set; }

        internal Func<string, string, string> OnBuildCaseWorkbookPath { get; set; }

        internal Func<string, bool> OnIsUnderSyncRoot { get; set; }

        internal Func<string, string> OnBuildLocalWorkingCaseWorkbookPath { get; set; }

        internal Func<string, string, bool> OnMoveLocalWorkingCaseToFinalPath { get; set; }

        internal string ResolveSystemRoot(Excel.Workbook kernelWorkbook)
        {
            return OnResolveSystemRoot == null ? string.Empty : OnResolveSystemRoot(kernelWorkbook);
        }

        internal string ResolveBaseWorkbookPath(string systemRoot)
        {
            return OnResolveBaseWorkbookPath == null ? string.Empty : OnResolveBaseWorkbookPath(systemRoot);
        }

        internal string ResolveCaseFolderPath(KernelCaseCreationRequest request, string folderName)
        {
            return OnResolveCaseFolderPath == null ? string.Empty : OnResolveCaseFolderPath(request, folderName);
        }

        internal bool EnsureFolderExists(string folderPath)
        {
            return OnEnsureFolderExists == null || OnEnsureFolderExists(folderPath);
        }

        internal string ResolveCaseWorkbookExtension(string baseWorkbookPath)
        {
            return OnResolveCaseWorkbookExtension == null ? string.Empty : OnResolveCaseWorkbookExtension(baseWorkbookPath);
        }

        internal string BuildCaseWorkbookPath(string folderPath, string caseWorkbookName)
        {
            return OnBuildCaseWorkbookPath == null ? string.Empty : OnBuildCaseWorkbookPath(folderPath, caseWorkbookName);
        }

        internal bool IsUnderSyncRoot(string path)
        {
            return OnIsUnderSyncRoot != null && OnIsUnderSyncRoot(path);
        }

        internal string BuildLocalWorkingCaseWorkbookPath(string finalCaseWorkbookPath)
        {
            return OnBuildLocalWorkingCaseWorkbookPath == null ? string.Empty : OnBuildLocalWorkingCaseWorkbookPath(finalCaseWorkbookPath);
        }

        internal bool MoveLocalWorkingCaseToFinalPath(string localWorkingPath, string finalCaseWorkbookPath)
        {
            return OnMoveLocalWorkingCaseToFinalPath == null
                || OnMoveLocalWorkingCaseToFinalPath(localWorkingPath, finalCaseWorkbookPath);
        }
    }

    internal sealed class DocumentExecutionEligibilityService
    {
        internal Func<Excel.Workbook, string, string, CaseInfoSystem.ExcelAddIn.Domain.DocumentExecutionEligibility> OnEvaluate { get; set; }

        internal CaseInfoSystem.ExcelAddIn.Domain.DocumentExecutionEligibility Evaluate(Excel.Workbook workbook, string actionKind, string key)
        {
            return OnEvaluate != null
                ? OnEvaluate(workbook, actionKind, key)
                : new CaseInfoSystem.ExcelAddIn.Domain.DocumentExecutionEligibility(false, string.Empty, null, null);
        }
    }

    internal sealed class DocumentExecutionPolicyService
    {
        internal Func<CaseInfoSystem.ExcelAddIn.Domain.DocumentTemplateSpec, bool> OnIsVstoExecutionAllowed { get; set; }

        internal Func<CaseInfoSystem.ExcelAddIn.Domain.DocumentTemplateSpec, bool> OnIsRolloutReady { get; set; }

        internal string AllowlistPath { get; set; } = string.Empty;

        internal bool IsVstoExecutionAllowed(CaseInfoSystem.ExcelAddIn.Domain.DocumentTemplateSpec templateSpec)
        {
            return OnIsVstoExecutionAllowed != null && OnIsVstoExecutionAllowed(templateSpec);
        }

        internal string GetAllowlistPath()
        {
            return AllowlistPath;
        }

        internal bool IsRolloutReady(CaseInfoSystem.ExcelAddIn.Domain.DocumentTemplateSpec templateSpec)
        {
            return OnIsRolloutReady != null && OnIsRolloutReady(templateSpec);
        }
    }

    internal sealed class DocumentCreateService
    {
        internal Action<Excel.Workbook, CaseInfoSystem.ExcelAddIn.Domain.DocumentTemplateSpec, CaseInfoSystem.ExcelAddIn.Domain.CaseContext> OnExecute { get; set; }

        internal void Execute(Excel.Workbook workbook, CaseInfoSystem.ExcelAddIn.Domain.DocumentTemplateSpec templateSpec, CaseInfoSystem.ExcelAddIn.Domain.CaseContext caseContext)
        {
            OnExecute?.Invoke(workbook, templateSpec, caseContext);
        }
    }

    internal sealed class AccountingSetCommandService
    {
        internal Action<Excel.Workbook> OnExecute { get; set; }

        internal void Execute(Excel.Workbook workbook)
        {
            OnExecute?.Invoke(workbook);
        }
    }

    internal sealed class CaseListRegistrationService
    {
        internal Func<Excel.Workbook, CaseInfoSystem.ExcelAddIn.Domain.CaseListRegistrationResult> OnExecute { get; set; }

        internal CaseInfoSystem.ExcelAddIn.Domain.CaseListRegistrationResult Execute(Excel.Workbook workbook)
        {
            return OnExecute != null
                ? OnExecute(workbook)
                : new CaseInfoSystem.ExcelAddIn.Domain.CaseListRegistrationResult();
        }
    }

    internal sealed class CaseContextFactory
    {
        internal Func<Excel.Workbook, CaseInfoSystem.ExcelAddIn.Domain.CaseContext> OnCreateForCaseListRegistration { get; set; }

        internal CaseInfoSystem.ExcelAddIn.Domain.CaseContext CreateForCaseListRegistration(Excel.Workbook caseWorkbook)
        {
            return OnCreateForCaseListRegistration == null ? null : OnCreateForCaseListRegistration(caseWorkbook);
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

    internal sealed class KernelWorkbookLifecycleService
    {
        internal IDisposable BeginManagedCloseScope(Excel.Workbook workbook) => new NoOpDisposable();

        internal bool RequestManagedCloseFromHomeExit(Excel.Workbook workbook) => true;

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    internal sealed class DocumentOutputService
    {
        internal DocumentOutputService(CaseInfoSystem.ExcelAddIn.Infrastructure.ExcelInteropService excelInteropService, CaseInfoSystem.ExcelAddIn.Infrastructure.PathCompatibilityService pathCompatibilityService, CaseInfoSystem.ExcelAddIn.Infrastructure.Logger logger)
        {
        }

        internal string PrepareSavePath(string rawFullPath) => rawFullPath ?? string.Empty;
    }
}

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    internal sealed class CaseListRegistrationResult
    {
        internal bool Success { get; set; }

        internal int RegisteredRow { get; set; }

        internal string Message { get; set; } = string.Empty;
    }

    internal enum DocumentExecutionMode
    {
        Disabled = 0,
        PilotOnly = 1,
        AllowlistedOnly = 2
    }

    internal sealed class DocumentExecutionEligibility
    {
        internal DocumentExecutionEligibility(bool canExecuteInVsto, string reason, DocumentTemplateSpec templateSpec, CaseContext caseContext)
        {
            CanExecuteInVsto = canExecuteInVsto;
            Reason = reason ?? string.Empty;
            TemplateSpec = templateSpec;
            CaseContext = caseContext;
        }

        internal bool CanExecuteInVsto { get; }

        internal string Reason { get; }

        internal DocumentTemplateSpec TemplateSpec { get; }

        internal CaseContext CaseContext { get; }
    }

    internal sealed class DocumentTemplateSpec
    {
        internal string TemplateFileName { get; set; } = string.Empty;
    }

    internal sealed class CaseContext
    {
        internal Excel.Workbook KernelWorkbook { get; set; }
    }
}

namespace CaseInfoSystem.ExcelAddIn.UI
{
    internal sealed class ExcelWindowOwner : IDisposable
    {
        internal static Func<Microsoft.Office.Interop.Excel.Window, ExcelWindowOwner> OnFrom { get; set; }

        internal bool Disposed { get; private set; }

        internal static ExcelWindowOwner From(Microsoft.Office.Interop.Excel.Window window)
        {
            return OnFrom == null ? new ExcelWindowOwner() : OnFrom(window);
        }

        public void Dispose()
        {
            Disposed = true;
        }
    }

    internal static class CompletionNoticeForm
    {
        internal static Action<ExcelWindowOwner, string, string> OnShowNotice { get; set; }

        internal static void ShowNotice(ExcelWindowOwner owner, string title, string message)
        {
            OnShowNotice?.Invoke(owner, title, message);
        }
    }
}
