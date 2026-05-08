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

        internal Func<Excel.Workbook, string, Excel.Worksheet> OnFindWorksheetByCodeName { get; set; }

        internal Func<Excel.Workbook, string, Excel.Worksheet> OnFindWorksheetByName { get; set; }

        internal Func<string, Excel.Workbook> OnFindOpenWorkbook { get; set; }

        internal Action<Excel.Workbook, string, string> OnSetDocumentProperty { get; set; }

        internal Func<Excel.Worksheet, IReadOnlyDictionary<string, string>> OnReadKeyValueMapFromColumnsAandB { get; set; }

        internal Func<Excel.Worksheet, IReadOnlyList<IReadOnlyDictionary<string, string>>> OnReadRecordsFromHeaderRow { get; set; }

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
            OnSetDocumentProperty?.Invoke(workbook, propertyName, value ?? string.Empty);
        }

        internal Excel.Window GetFirstVisibleWindow(Excel.Workbook workbook) => workbook == null ? null : workbook.Windows.FirstOrDefault(window => window.Visible);

        internal string GetActiveSheetCodeName(Excel.Workbook workbook) => workbook?.ActiveSheet?.CodeName ?? string.Empty;

        internal Excel.Worksheet FindWorksheetByCodeName(Excel.Workbook workbook, string sheetCodeName)
        {
            if (OnFindWorksheetByCodeName != null)
            {
                return OnFindWorksheetByCodeName(workbook, sheetCodeName);
            }

            return workbook?.Worksheets.FirstOrDefault(worksheet => string.Equals(worksheet?.CodeName, sheetCodeName, StringComparison.OrdinalIgnoreCase));
        }

        internal Excel.Worksheet FindWorksheetByName(Excel.Workbook workbook, string sheetName)
        {
            if (OnFindWorksheetByName != null)
            {
                return OnFindWorksheetByName(workbook, sheetName);
            }

            return workbook?.Worksheets.FirstOrDefault(worksheet => string.Equals(worksheet?.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        }

        internal IReadOnlyDictionary<string, string> ReadKeyValueMapFromColumnsAandB(Excel.Worksheet worksheet)
        {
            return OnReadKeyValueMapFromColumnsAandB == null
                ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                : OnReadKeyValueMapFromColumnsAandB(worksheet) ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        internal IReadOnlyList<IReadOnlyDictionary<string, string>> ReadRecordsFromHeaderRow(Excel.Worksheet worksheet)
        {
            return OnReadRecordsFromHeaderRow == null
                ? Array.Empty<IReadOnlyDictionary<string, string>>()
                : OnReadRecordsFromHeaderRow(worksheet) ?? Array.Empty<IReadOnlyDictionary<string, string>>();
        }

        internal bool TryNormalizeCaseListRowHeight(CaseInfoSystem.ExcelAddIn.Domain.CaseContext context)
        {
            return OnTryNormalizeCaseListRowHeight == null || OnTryNormalizeCaseListRowHeight(context);
        }

        internal Excel.Workbook FindOpenWorkbook(string workbookPath)
        {
            return OnFindOpenWorkbook == null ? null : OnFindOpenWorkbook(workbookPath);
        }

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

        internal void RegisterKnownCasePath(string path)
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

        internal bool HideApplicationWindow(string reason, string workbookFullName)
        {
            if (_application != null)
            {
                _application.Visible = false;
            }

            return true;
        }

        internal bool ShowApplicationWindow(string reason, string workbookFullName)
        {
            if (_application != null)
            {
                _application.Visible = true;
            }

            return true;
        }

        internal bool TryBringApplicationToForeground(string reason, string workbookFullName) => true;

        internal bool TryRestoreMainWindow(bool bringToFront) => true;

        internal bool TryRestoreWorkbookWindow(Excel.Workbook workbook, bool bringToFront) => true;

        internal bool TryRecoverWorkbookWindow(Excel.Workbook workbook, string reason, bool bringToFront)
        {
            if (workbook == null)
            {
                return false;
            }

            EnsureApplicationVisible(reason, workbook.FullName);

            Excel.Window window = _excelInteropService == null
                ? null
                : _excelInteropService.GetFirstVisibleWindow(workbook);
            if (window == null)
            {
                workbook.Activate();
                if (workbook.Windows.Count > 0)
                {
                    window = workbook.Windows[1];
                }
            }

            if (window == null)
            {
                window = workbook.NewWindow();
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

    internal sealed class MasterTemplateCatalogService
    {
        internal Action<Excel.Workbook> OnInvalidateCache { get; set; }

        internal void InvalidateCache(Excel.Workbook workbook)
        {
            OnInvalidateCache?.Invoke(workbook);
        }
    }

    internal sealed class UserErrorService
    {
        internal Action<string, Exception> OnShowUserError { get; set; }

        internal void ShowUserError(string context, Exception ex)
        {
            OnShowUserError?.Invoke(context, ex);
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

    internal sealed class TaskPaneSnapshotBuilderService : CaseInfoSystem.ExcelAddIn.App.ICaseTaskPaneSnapshotReader
    {
        internal enum TaskPaneSnapshotSource
        {
            None = 0,
            CaseCache = 1,
            BaseCache = 2,
            BaseCacheFallback = 3,
            MasterListRebuild = 4,
        }

        internal sealed class TaskPaneBuildResult
        {
            internal TaskPaneBuildResult(string snapshotText, bool updatedCaseSnapshotCache)
                : this(
                    snapshotText,
                    updatedCaseSnapshotCache,
                    TaskPaneSnapshotSource.None,
                    string.Empty,
                    masterListRebuildAttempted: false,
                    masterListRebuildSucceeded: false,
                    failureReason: string.Empty,
                    degradedReason: string.Empty)
            {
            }

            internal TaskPaneBuildResult(
                string snapshotText,
                bool updatedCaseSnapshotCache,
                TaskPaneSnapshotSource snapshotSource,
                string fallbackReasons,
                bool masterListRebuildAttempted,
                bool masterListRebuildSucceeded,
                string failureReason,
                string degradedReason)
            {
                SnapshotText = snapshotText ?? string.Empty;
                UpdatedCaseSnapshotCache = updatedCaseSnapshotCache;
                SnapshotSource = snapshotSource;
                FallbackReasons = fallbackReasons ?? string.Empty;
                MasterListRebuildAttempted = masterListRebuildAttempted;
                MasterListRebuildSucceeded = masterListRebuildSucceeded;
                FailureReason = failureReason ?? string.Empty;
                DegradedReason = degradedReason ?? string.Empty;
            }

            internal string SnapshotText { get; }

            internal bool UpdatedCaseSnapshotCache { get; }

            internal TaskPaneSnapshotSource SnapshotSource { get; }

            internal string FallbackReasons { get; }

            internal bool MasterListRebuildAttempted { get; }

            internal bool MasterListRebuildSucceeded { get; }

            internal bool SnapshotTextAvailable
            {
                get { return !string.IsNullOrWhiteSpace(SnapshotText); }
            }

            internal string FailureReason { get; }

            internal string DegradedReason { get; }
        }

        internal Func<Excel.Workbook, TaskPaneBuildResult> OnBuildSnapshotText { get; set; }

        internal TaskPaneBuildResult BuildSnapshotText(Excel.Workbook workbook)
        {
            return OnBuildSnapshotText != null
                ? OnBuildSnapshotText(workbook)
                : new TaskPaneBuildResult(string.Empty, updatedCaseSnapshotCache: false);
        }

        TaskPaneBuildResult CaseInfoSystem.ExcelAddIn.App.ICaseTaskPaneSnapshotReader.BuildSnapshotText(Excel.Workbook workbook)
        {
            return BuildSnapshotText(workbook);
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

        internal Func<Excel.Workbook, CaseInfoSystem.ExcelAddIn.Domain.CaseContext> OnCreateForDocumentCreate { get; set; }

        internal CaseInfoSystem.ExcelAddIn.Domain.CaseContext CreateForCaseListRegistration(Excel.Workbook caseWorkbook)
        {
            return OnCreateForCaseListRegistration == null ? null : OnCreateForCaseListRegistration(caseWorkbook);
        }

        internal CaseInfoSystem.ExcelAddIn.Domain.CaseContext CreateForDocumentCreate(Excel.Workbook caseWorkbook)
        {
            return OnCreateForDocumentCreate == null ? null : OnCreateForDocumentCreate(caseWorkbook);
        }
    }

    internal sealed class DocumentNamePromptService
    {
        internal Func<Excel.Workbook, string, bool> OnTryPrepare { get; set; }

        internal Func<Excel.Workbook, string, DocumentNameOverrideScope> OnCreateScope { get; set; }

        internal bool TryPrepare(Excel.Workbook workbook, string key, out DocumentNameOverrideScope scope)
        {
            scope = OnCreateScope == null ? null : OnCreateScope(workbook, key);
            return OnTryPrepare == null || OnTryPrepare(workbook, key);
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
        private Action<string, Excel.Workbook, bool> _homeManagedCloseStarted;
        private Action<string, Excel.Workbook, bool> _homeManagedCloseSucceeded;
        private Action<string, Excel.Workbook, bool, Exception> _homeManagedCloseFailed;

        internal IDisposable BeginManagedCloseScope(Excel.Workbook workbook) => new NoOpDisposable();

        internal bool RequestManagedCloseFromHomeExit(Excel.Workbook workbook) => true;

        internal void RegisterHomeManagedCloseCallbacks(
            Action<string, Excel.Workbook, bool> onStarted,
            Action<string, Excel.Workbook, bool> onSucceeded,
            Action<string, Excel.Workbook, bool, Exception> onFailed)
        {
            _homeManagedCloseStarted = onStarted;
            _homeManagedCloseSucceeded = onSucceeded;
            _homeManagedCloseFailed = onFailed;
        }

        internal void SimulateManagedCloseSuccess(Excel.Workbook workbook, bool saveChanges = false)
        {
            string workbookKey = workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
            _homeManagedCloseStarted?.Invoke(workbookKey, workbook, saveChanges);
            _homeManagedCloseSucceeded?.Invoke(workbookKey, workbook, saveChanges);
        }

        internal void SimulateManagedCloseFailure(Excel.Workbook workbook, Exception exception, bool saveChanges = false)
        {
            string workbookKey = workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
            _homeManagedCloseStarted?.Invoke(workbookKey, workbook, saveChanges);
            _homeManagedCloseFailed?.Invoke(workbookKey, workbook, saveChanges, exception);
        }

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    internal sealed class DocumentOutputService
    {
        internal Func<Excel.Workbook, string> OnResolveWorkbookFolder { get; set; }

        internal Func<Excel.Workbook, string, string, string> OnBuildOutputFileName { get; set; }

        internal DocumentOutputService(CaseInfoSystem.ExcelAddIn.Infrastructure.ExcelInteropService excelInteropService, CaseInfoSystem.ExcelAddIn.Infrastructure.PathCompatibilityService pathCompatibilityService, CaseInfoSystem.ExcelAddIn.Infrastructure.Logger logger)
        {
        }

        internal string PrepareSavePath(string rawFullPath) => rawFullPath ?? string.Empty;

        internal string ResolveWorkbookFolder(Excel.Workbook workbook)
        {
            return OnResolveWorkbookFolder == null ? string.Empty : OnResolveWorkbookFolder(workbook) ?? string.Empty;
        }

        internal string BuildOutputFileName(Excel.Workbook workbook, string documentName, string customerName)
        {
            return OnBuildOutputFileName == null ? string.Empty : OnBuildOutputFileName(workbook, documentName, customerName) ?? string.Empty;
        }
    }

    internal sealed class AccountingSetNamingService
    {
        internal Func<Excel.Workbook, string, string, string, string> OnBuildCaseOutputPath { get; set; }

        internal string BuildCaseOutputPath(Excel.Workbook workbook, string outputFolderPath, string customerName, string templatePath)
        {
            return OnBuildCaseOutputPath == null
                ? string.Empty
                : OnBuildCaseOutputPath(workbook, outputFolderPath, customerName, templatePath) ?? string.Empty;
        }
    }

    internal sealed class AccountingTemplateResolver
    {
        internal Func<Excel.Workbook, string> OnResolveTemplatePath { get; set; }

        internal string ResolveTemplatePath(Excel.Workbook workbook)
        {
            return OnResolveTemplatePath == null ? string.Empty : OnResolveTemplatePath(workbook) ?? string.Empty;
        }
    }

    internal sealed class AccountingWorkbookService
    {
        internal Func<string, Excel.Workbook> OnOpenInCurrentApplication { get; set; }

        internal Action<Excel.Workbook, bool> OnSetWorkbookWindowsVisible { get; set; }

        internal Func<IDisposable> OnBeginInitializationScope { get; set; }

        internal Action<Excel.Workbook, IEnumerable<string>, string, string> OnWriteSameValueToSheets { get; set; }

        internal Action<Excel.Workbook, string, string, string> OnWriteCell { get; set; }

        internal Action<Excel.Worksheet, string, object> OnWriteCellValue { get; set; }

        internal Action<Excel.Worksheet, string, object[,]> OnWriteRangeValues { get; set; }

        internal Func<Excel.Workbook, string, AccountingLawyerMappingResult> OnReflectLawyers { get; set; }

        internal Action<Excel.Workbook> OnActivateInvoiceEntry { get; set; }

        internal Excel.Workbook OpenInCurrentApplication(string workbookPath)
        {
            return OnOpenInCurrentApplication == null ? null : OnOpenInCurrentApplication(workbookPath);
        }

        internal void SetWorkbookWindowsVisible(Excel.Workbook workbook, bool visible)
        {
            OnSetWorkbookWindowsVisible?.Invoke(workbook, visible);
        }

        internal IDisposable BeginInitializationScope()
        {
            return OnBeginInitializationScope == null ? new NoOpDisposable() : OnBeginInitializationScope() ?? new NoOpDisposable();
        }

        internal void WriteSameValueToSheets(Excel.Workbook workbook, IEnumerable<string> sheetNames, string address, string valueText)
        {
            OnWriteSameValueToSheets?.Invoke(workbook, sheetNames, address, valueText);
        }

        internal void WriteCell(Excel.Workbook workbook, string sheetName, string address, string valueText)
        {
            OnWriteCell?.Invoke(workbook, sheetName, address, valueText);
        }

        internal void WriteCellValue(Excel.Worksheet worksheet, string address, object value)
        {
            OnWriteCellValue?.Invoke(worksheet, address, value);
        }

        internal void WriteRangeValues(Excel.Worksheet worksheet, string address, object[,] values)
        {
            OnWriteRangeValues?.Invoke(worksheet, address, values);
        }

        internal AccountingLawyerMappingResult ReflectLawyers(Excel.Workbook workbook, string lawyerLinesText)
        {
            return OnReflectLawyers == null ? new AccountingLawyerMappingResult() : OnReflectLawyers(workbook, lawyerLinesText);
        }

        internal void ActivateInvoiceEntry(Excel.Workbook workbook)
        {
            OnActivateInvoiceEntry?.Invoke(workbook);
        }

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }

    internal sealed class AccountingLawyerMappingResult
    {
        internal int AssignedCount { get; set; }

        internal int MissingMatchCount { get; set; }

        internal int OverflowCount { get; set; }
    }

    internal sealed class AccountingSetPresentationWaitService
    {
        internal const string CreatingStageTitle = "Creating";

        internal const string OpeningWorkbookStageTitle = "Opening";

        internal const string ApplyingInitialDataStageTitle = "Applying";

        internal const string ShowingInputScreenStageTitle = "Showing";

        internal Func<System.Diagnostics.Stopwatch, WaitSession> OnShowWaiting { get; set; }

        internal WaitSession ShowWaiting(System.Diagnostics.Stopwatch commandStopwatch)
        {
            return OnShowWaiting == null ? new WaitSession() : OnShowWaiting(commandStopwatch) ?? new WaitSession();
        }

        internal sealed class WaitSession : IDisposable
        {
            internal readonly List<string> Stages = new List<string>();

            internal Action OnClose { get; set; }

            internal Action OnDispose { get; set; }

            internal void UpdateStage(string title, string detail = null)
            {
                Stages.Add(title ?? string.Empty);
            }

            internal void Close()
            {
                OnClose?.Invoke();
            }

            public void Dispose()
            {
                OnDispose?.Invoke();
            }
        }
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
        WarmupEnabledProfileA = 1,
        WarmupEnabledProfileB = 2
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
        internal Excel.Workbook CaseWorkbook { get; set; }

        internal Excel.Workbook KernelWorkbook { get; set; }

        internal Excel.Worksheet CaseListWorksheet { get; set; }

        internal IReadOnlyDictionary<string, string> CaseValues { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        internal string CustomerName { get; set; } = string.Empty;

        internal string WorkbookName { get; set; } = string.Empty;

        internal string WorkbookPath { get; set; } = string.Empty;

        internal string HomeSheetName { get; set; } = string.Empty;

        internal string SystemRoot { get; set; } = string.Empty;
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
