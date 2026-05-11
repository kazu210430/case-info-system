using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookCloseService
    {
        private readonly Excel.Application _application;
        private readonly KernelCaseInteractionState _kernelCaseInteractionState;
        private readonly Logger _logger;
        private readonly KernelWorkbookBindingService _bindingService;
        private readonly KernelWorkbookDisplayService _displayService;
        private readonly TestHooks _testHooks;
        private readonly KernelHomeSessionCloseCoordinator _homeSessionCloseCoordinator;
        private KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;
        private PendingHomeSessionClose _pendingHomeSessionClose;
        private Action _homeSessionCloseReadyToCloseForm;
        private Action _homeSessionCloseFailed;

        internal KernelWorkbookCloseService(
            Excel.Application application,
            KernelCaseInteractionState kernelCaseInteractionState,
            Logger logger,
            KernelWorkbookBindingService bindingService,
            KernelWorkbookDisplayService displayService,
            TestHooks testHooks = null)
        {
            _application = application;
            _kernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _bindingService = bindingService ?? throw new ArgumentNullException(nameof(bindingService));
            _displayService = displayService ?? throw new ArgumentNullException(nameof(displayService));
            _testHooks = testHooks;
            _homeSessionCloseCoordinator = new KernelHomeSessionCloseCoordinator(this);
        }

        internal void SetLifecycleService(KernelWorkbookLifecycleService kernelWorkbookLifecycleService)
        {
            _kernelWorkbookLifecycleService = kernelWorkbookLifecycleService ?? throw new ArgumentNullException(nameof(kernelWorkbookLifecycleService));
            _kernelWorkbookLifecycleService.RegisterHomeManagedCloseCallbacks(
                HandleManagedHomeSessionCloseStarted,
                HandleManagedHomeSessionCloseSucceeded,
                HandleManagedHomeSessionCloseFailed);
        }

        internal void RegisterHomeSessionCloseObserver(Action onReadyToCloseForm, Action onFailed)
        {
            _homeSessionCloseReadyToCloseForm = onReadyToCloseForm;
            _homeSessionCloseFailed = onFailed;
        }

        internal KernelHomeSessionCloseRequestStatus RequestCloseHomeSessionFromForm(bool saveKernelWorkbook, string entryPoint)
        {
            string caller = ResolveExternalCaller();
            return _homeSessionCloseCoordinator.ExecuteForForm(saveKernelWorkbook, entryPoint, caller);
        }

        internal void FinalizePendingHomeSessionCloseAfterFormClosed()
        {
            _homeSessionCloseCoordinator.FinalizePendingCloseAfterFormClosed();
        }

        internal void CloseHomeSession()
        {
            CloseHomeSession(saveKernelWorkbook: false, entryPoint: "CloseHomeSession");
        }

        internal void CloseHomeSessionSavingKernel()
        {
            CloseHomeSession(saveKernelWorkbook: true, entryPoint: "CloseHomeSessionSavingKernel");
        }

        internal void CloseKernelWorkbookWithoutLifecycleCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.CloseKernelWorkbookWithoutLifecycle != null)
            {
                _testHooks.CloseKernelWorkbookWithoutLifecycle(workbook);
                return;
            }

            CloseWorkbookWithoutSave(workbook);
        }

        internal void SaveAndCloseKernelWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            string workbookFullName = _bindingService.GetWorkbookFullName(workbook);
            bool requiresSave = RequiresSave(workbook);
            _logger.Info(
                "SaveAndCloseKernelWorkbook started. workbook="
                + workbookFullName
                + ", requiresSave="
                + requiresSave.ToString());

            if (requiresSave)
            {
                workbook.Save();
                _logger.Info("SaveAndCloseKernelWorkbook saved workbook=" + workbookFullName);
            }
            else
            {
                _logger.Info("SaveAndCloseKernelWorkbook skipped save because workbook was already saved. workbook=" + workbookFullName);
            }

            CloseWorkbookWithoutSave(workbook);

            _logger.Info("SaveAndCloseKernelWorkbook closed workbook=" + workbookFullName);
        }

        internal void QuitApplicationCore()
        {
            if (_testHooks != null && _testHooks.QuitApplication != null)
            {
                _testHooks.QuitApplication();
                return;
            }

            bool previousDisplayAlerts = true;
            bool hasDisplayAlertsSnapshot = false;
            try
            {
                previousDisplayAlerts = _application.DisplayAlerts;
                hasDisplayAlertsSnapshot = true;
                _application.DisplayAlerts = false;
                _application.Quit();
            }
            catch
            {
                if (hasDisplayAlertsSnapshot)
                {
                    try
                    {
                        _application.DisplayAlerts = previousDisplayAlerts;
                    }
                    catch
                    {
                    }
                }

                throw;
            }
        }

        private void HandleManagedHomeSessionCloseStarted(string workbookKey, Excel.Workbook workbook, bool saveChanges)
        {
            PendingHomeSessionClose pendingClose = FindPendingHomeSessionClose(workbookKey);
            if (pendingClose == null)
            {
                return;
            }

            _logger.Info(
                "CloseHomeSession managed close started. entryPoint="
                + (pendingClose.EntryPoint ?? string.Empty)
                + ", workbook="
                + _bindingService.GetWorkbookFullName(workbook)
                + ", saveChanges="
                + saveChanges.ToString());
        }

        private void HandleManagedHomeSessionCloseSucceeded(string workbookKey, Excel.Workbook workbook, bool saveChanges)
        {
            PendingHomeSessionClose pendingClose = FindPendingHomeSessionClose(workbookKey);
            if (pendingClose == null)
            {
                return;
            }

            _logger.Info(
                "CloseHomeSession managed close succeeded. entryPoint="
                + (pendingClose.EntryPoint ?? string.Empty)
                + ", workbook="
                + _bindingService.GetWorkbookFullName(workbook)
                + ", saveChanges="
                + saveChanges.ToString());

            if (pendingClose.DeferSessionCompletionUntilFormClosed)
            {
                pendingClose.MarkBackendCloseCompleted();
                NotifyHomeSessionCloseReadyToCloseForm();
                return;
            }

            _homeSessionCloseCoordinator.CompleteDeferredHomeSession(TakePendingHomeSessionClose(workbookKey), workbook);
        }

        private void HandleManagedHomeSessionCloseFailed(string workbookKey, Excel.Workbook workbook, bool saveChanges, Exception exception)
        {
            PendingHomeSessionClose pendingClose = TakePendingHomeSessionClose(workbookKey);
            if (pendingClose == null)
            {
                return;
            }

            _logger.Error(
                "CloseHomeSession managed close failed. HOME display state and binding were preserved. entryPoint="
                + (pendingClose.EntryPoint ?? string.Empty)
                + ", workbook="
                + _bindingService.GetWorkbookFullName(workbook)
                + ", saveChanges="
                + saveChanges.ToString()
                + ", exceptionType="
                + (exception == null ? string.Empty : exception.GetType().FullName ?? string.Empty)
                + ", exceptionMessage="
                + (exception == null ? string.Empty : exception.Message ?? string.Empty),
                exception);

            if (pendingClose.DeferSessionCompletionUntilFormClosed)
            {
                NotifyHomeSessionCloseFailed();
            }
        }

        private void CloseHomeSession(bool saveKernelWorkbook, string entryPoint)
        {
            string caller = ResolveExternalCaller();
            _homeSessionCloseCoordinator.Execute(saveKernelWorkbook, entryPoint, caller);
        }

        private void NotifyHomeSessionCloseReadyToCloseForm()
        {
            if (_homeSessionCloseReadyToCloseForm == null)
            {
                return;
            }

            _homeSessionCloseReadyToCloseForm();
        }

        private void NotifyHomeSessionCloseFailed()
        {
            if (_homeSessionCloseFailed == null)
            {
                return;
            }

            _homeSessionCloseFailed();
        }

        private void RegisterPendingHomeSessionClose(
            string workbookKey,
            bool saveKernelWorkbook,
            KernelHomeSessionCompletionAction completionAction,
            string entryPoint,
            bool deferSessionCompletionUntilFormClosed,
            bool backendCloseCompleted)
        {
            _pendingHomeSessionClose = new PendingHomeSessionClose(
                workbookKey,
                saveKernelWorkbook,
                completionAction,
                entryPoint,
                deferSessionCompletionUntilFormClosed,
                backendCloseCompleted);
            _logger.Info(
                "CloseHomeSession pending close registered. entryPoint="
                + (entryPoint ?? string.Empty)
                + ", workbook="
                + (workbookKey ?? string.Empty)
                + ", completionAction="
                + completionAction.ToString()
                + ", deferSessionCompletionUntilFormClosed="
                + deferSessionCompletionUntilFormClosed.ToString()
                + ", backendCloseCompleted="
                + backendCloseCompleted.ToString());
        }

        private PendingHomeSessionClose FindPendingHomeSessionClose(string workbookKey)
        {
            if (_pendingHomeSessionClose == null
                || string.IsNullOrWhiteSpace(workbookKey)
                || !string.Equals(_pendingHomeSessionClose.WorkbookKey, workbookKey, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return _pendingHomeSessionClose;
        }

        private PendingHomeSessionClose TakePendingHomeSessionClose(string workbookKey = null)
        {
            PendingHomeSessionClose pendingClose = string.IsNullOrWhiteSpace(workbookKey)
                ? _pendingHomeSessionClose
                : FindPendingHomeSessionClose(workbookKey);
            if (pendingClose != null)
            {
                _pendingHomeSessionClose = null;
            }

            return pendingClose;
        }

        private bool RequestManagedCloseFromHomeExitCore(Excel.Workbook workbook)
        {
            return _testHooks != null && _testHooks.RequestManagedCloseFromHomeExit != null
                ? _testHooks.RequestManagedCloseFromHomeExit(workbook)
                : _kernelWorkbookLifecycleService.RequestManagedCloseFromHomeExit(workbook);
        }

        private void SaveAndCloseKernelWorkbookCore(Excel.Workbook workbook)
        {
            if (_testHooks != null && _testHooks.SaveAndCloseKernelWorkbook != null)
            {
                _testHooks.SaveAndCloseKernelWorkbook(workbook);
                return;
            }

            SaveAndCloseKernelWorkbook(workbook);
        }

        private void CloseWorkbookWithoutSave(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            bool previousDisplayAlerts = true;
            bool hasDisplayAlertsSnapshot = false;
            try
            {
                previousDisplayAlerts = _application.DisplayAlerts;
                hasDisplayAlertsSnapshot = true;
                _application.DisplayAlerts = false;
                WorkbookCloseInteropHelper.CloseWithoutSave(
                    workbook,
                    _logger,
                    "KernelWorkbookCloseService.CloseKernelWorkbookWithoutLifecycleCore");
            }
            finally
            {
                if (hasDisplayAlertsSnapshot)
                {
                    _application.DisplayAlerts = previousDisplayAlerts;
                }
            }
        }

        private static bool RequiresSave(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                return !workbook.Saved;
            }
            catch
            {
                return true;
            }
        }

        private string ResolveExternalCaller()
        {
            try
            {
                var stackTrace = new System.Diagnostics.StackTrace(skipFrames: 1, fNeedFileInfo: false);
                System.Diagnostics.StackFrame[] frames = stackTrace.GetFrames();
                if (frames == null)
                {
                    return string.Empty;
                }

                foreach (System.Diagnostics.StackFrame frame in frames)
                {
                    var method = frame.GetMethod();
                    if (method == null)
                    {
                        continue;
                    }

                    Type declaringType = method.DeclaringType;
                    if (declaringType == typeof(KernelWorkbookCloseService)
                        || declaringType == typeof(KernelWorkbookService)
                        || declaringType == typeof(KernelHomeSessionCloseCoordinator))
                    {
                        continue;
                    }

                    string typeName = declaringType == null ? string.Empty : declaringType.FullName ?? string.Empty;
                    return string.IsNullOrWhiteSpace(typeName) ? method.Name : typeName + "." + method.Name;
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private sealed class KernelHomeSessionCloseCoordinator
        {
            private readonly KernelWorkbookCloseService _owner;

            internal KernelHomeSessionCloseCoordinator(KernelWorkbookCloseService owner)
            {
                _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            }

            internal void Execute(bool saveKernelWorkbook, string entryPoint, string caller)
            {
                ExecuteCore(
                    saveKernelWorkbook,
                    entryPoint,
                    caller,
                    deferSessionCompletionUntilFormClosed: false);
            }

            internal KernelHomeSessionCloseRequestStatus ExecuteForForm(bool saveKernelWorkbook, string entryPoint, string caller)
            {
                return ExecuteCore(
                    saveKernelWorkbook,
                    entryPoint,
                    caller,
                    deferSessionCompletionUntilFormClosed: true);
            }

            internal void FinalizePendingCloseAfterFormClosed()
            {
                PendingHomeSessionClose pendingClose = _owner.TakePendingHomeSessionClose();
                if (pendingClose == null)
                {
                    return;
                }

                if (!pendingClose.DeferSessionCompletionUntilFormClosed || !pendingClose.BackendCloseCompleted)
                {
                    _owner._logger.Warn(
                        "CloseHomeSession finalization skipped because pending close was not ready. entryPoint="
                        + (pendingClose.EntryPoint ?? string.Empty)
                        + ", backendCloseCompleted="
                        + pendingClose.BackendCloseCompleted.ToString());
                    return;
                }

                CompleteDeferredHomeSession(
                    pendingClose,
                    _owner._bindingService.ResolveWorkbookForHomeDisplayOrClose("FinalizePendingHomeSessionCloseAfterFormClosed"));
            }

            private KernelHomeSessionCloseRequestStatus ExecuteCore(
                bool saveKernelWorkbook,
                string entryPoint,
                string caller,
                bool deferSessionCompletionUntilFormClosed)
            {
                Excel.Workbook workbook = _owner._bindingService.ResolveWorkbookForHomeDisplayOrClose(entryPoint);
                bool otherVisibleWorkbookExists = _owner._bindingService.HasOtherVisibleWorkbook(workbook);
                bool otherWorkbookExists = _owner._bindingService.HasOtherWorkbook(workbook);
                bool skipDisplayRestoreForCaseCreation = KernelHomeSessionDisplayPolicy.ShouldSkipDisplayRestoreForCaseCreation(
                    saveKernelWorkbook,
                    _owner._kernelCaseInteractionState.IsKernelCaseCreationFlowActive,
                    otherVisibleWorkbookExists,
                    otherWorkbookExists);
                KernelHomeSessionCompletionAction completionAction = KernelHomeSessionDisplayPolicy.DecideCompletionAction(
                    skipDisplayRestoreForCaseCreation,
                    otherVisibleWorkbookExists,
                    otherWorkbookExists);
                _owner._displayService.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-enter entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", caller="
                    + caller
                    + ", saveKernelWorkbook="
                    + saveKernelWorkbook.ToString()
                    + ", workbook="
                    + _owner._displayService.FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _owner._displayService.FormatActiveExcelState()
                    + ", otherVisibleWorkbookExists="
                    + otherVisibleWorkbookExists.ToString()
                    + ", otherWorkbookExists="
                    + otherWorkbookExists.ToString()
                    + ", skipDisplayRestoreForCaseCreation="
                    + skipDisplayRestoreForCaseCreation.ToString()
                    + ", completionAction="
                    + completionAction.ToString()
                    + ", otherVisibleTargets="
                    + _owner._displayService.DescribeVisibleOtherWorkbookWindows(workbook));
                _owner._logger.Info(
                    "CloseHomeSession started. saveKernelWorkbook="
                    + saveKernelWorkbook.ToString()
                    + ", workbook="
                    + _owner._bindingService.GetWorkbookFullName(workbook)
                    + ", otherVisibleWorkbookExists="
                    + otherVisibleWorkbookExists.ToString()
                    + ", otherWorkbookExists="
                    + otherWorkbookExists.ToString()
                    + ", skipDisplayRestoreForCaseCreation="
                    + skipDisplayRestoreForCaseCreation.ToString());

                bool deferCompletion = false;
                try
                {
                    if (workbook != null
                        && !ExecuteCloseBranch(
                            workbook,
                            saveKernelWorkbook,
                            skipDisplayRestoreForCaseCreation,
                            completionAction,
                            entryPoint,
                            deferSessionCompletionUntilFormClosed,
                            out deferCompletion))
                    {
                        return KernelHomeSessionCloseRequestStatus.Rejected;
                    }
                }
                catch (Exception ex)
                {
                    _owner._logger.Error(
                        "CloseHomeSession failed before completion. entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", workbook="
                        + _owner._bindingService.GetWorkbookFullName(workbook)
                        + ", saveKernelWorkbook="
                        + saveKernelWorkbook.ToString(),
                        ex);
                    if (!deferSessionCompletionUntilFormClosed)
                    {
                        throw;
                    }

                    MessageBox.Show(
                        "保存または終了に失敗しました。もう一度お試しください。",
                        "案件情報System",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return KernelHomeSessionCloseRequestStatus.Rejected;
                }

                if (deferCompletion)
                {
                    return KernelHomeSessionCloseRequestStatus.Pending;
                }

                if (deferSessionCompletionUntilFormClosed)
                {
                    _owner.RegisterPendingHomeSessionClose(
                        _owner._bindingService.GetWorkbookFullName(workbook),
                        saveKernelWorkbook,
                        completionAction,
                        entryPoint,
                        deferSessionCompletionUntilFormClosed: true,
                        backendCloseCompleted: true);
                    return KernelHomeSessionCloseRequestStatus.Completed;
                }

                CompleteHomeSession(saveKernelWorkbook, completionAction, workbook, entryPoint);
                return KernelHomeSessionCloseRequestStatus.Completed;
            }

            private bool ExecuteCloseBranch(
                Excel.Workbook workbook,
                bool saveKernelWorkbook,
                bool skipDisplayRestoreForCaseCreation,
                KernelHomeSessionCompletionAction completionAction,
                string entryPoint,
                bool deferSessionCompletionUntilFormClosed,
                out bool deferCompletion)
            {
                deferCompletion = false;
                if (saveKernelWorkbook)
                {
                    _owner._displayService.LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=save-and-close, workbook="
                        + _owner._displayService.FormatWorkbookDescriptor(workbook)
                        + ", skipDisplayRestoreForCaseCreation="
                        + skipDisplayRestoreForCaseCreation.ToString());
                    if (skipDisplayRestoreForCaseCreation)
                    {
                        _owner._displayService.ConcealKernelWorkbookWindowsForCaseCreationCloseCore(workbook);
                    }

                    _owner.SaveAndCloseKernelWorkbookCore(workbook);
                    return true;
                }

                if (_owner._kernelWorkbookLifecycleService != null)
                {
                    _owner._displayService.LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=request-managed-close, workbook="
                        + _owner._displayService.FormatWorkbookDescriptor(workbook)
                        + ", lifecycleAvailable=True, activeState="
                        + _owner._displayService.FormatActiveExcelState());
                    bool closeScheduled = _owner.RequestManagedCloseFromHomeExitCore(workbook);
                    _owner._displayService.LogKernelFlickerTrace(
                        "source=KernelWorkbookService action=close-home-session-branch-result entryPoint="
                        + (entryPoint ?? string.Empty)
                        + ", branch=request-managed-close, workbook="
                        + _owner._displayService.FormatWorkbookDescriptor(workbook)
                        + ", closeScheduled="
                        + closeScheduled.ToString());
                    if (!closeScheduled)
                    {
                        _owner._displayService.LogKernelFlickerTrace(
                            "source=KernelWorkbookService action=close-home-session-end entryPoint="
                            + (entryPoint ?? string.Empty)
                            + ", result=canceled-before-managed-close, workbook="
                            + _owner._displayService.FormatWorkbookDescriptor(workbook));
                        _owner._logger.Info("CloseHomeSession canceled before managed close was scheduled.");
                        return false;
                    }

                    _owner.RegisterPendingHomeSessionClose(
                        _owner._bindingService.GetWorkbookFullName(workbook),
                        saveKernelWorkbook,
                        completionAction,
                        entryPoint,
                        deferSessionCompletionUntilFormClosed,
                        backendCloseCompleted: false);
                    deferCompletion = true;
                    return true;
                }

                _owner._displayService.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-branch entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", branch=close-without-lifecycle, workbook="
                    + _owner._displayService.FormatWorkbookDescriptor(workbook)
                    + ", lifecycleAvailable=False");
                _owner.CloseKernelWorkbookWithoutLifecycleCore(workbook);
                return true;
            }

            internal void CompleteDeferredHomeSession(PendingHomeSessionClose pendingClose, Excel.Workbook workbook)
            {
                if (pendingClose == null)
                {
                    return;
                }

                CompleteHomeSession(
                    pendingClose.SaveKernelWorkbook,
                    pendingClose.CompletionAction,
                    workbook,
                    pendingClose.EntryPoint);
            }

            private void CompleteHomeSession(
                bool saveKernelWorkbook,
                KernelHomeSessionCompletionAction completionAction,
                Excel.Workbook workbook,
                string entryPoint)
            {
                _owner._logger.Info(
                    "CloseHomeSession UI release executing. entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", completionAction="
                    + completionAction.ToString()
                    + ", workbook="
                    + _owner._bindingService.GetWorkbookFullName(workbook));
                _owner._displayService.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-completion entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", completionAction="
                    + completionAction.ToString()
                    + ", workbook="
                    + _owner._displayService.FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _owner._displayService.FormatActiveExcelState());
                if (completionAction == KernelHomeSessionCompletionAction.ReleaseHomeDisplayWithoutShowingExcelAndQuit)
                {
                    _owner._displayService.ReleaseHomeDisplayCore(false);
                    if (saveKernelWorkbook || _owner._kernelWorkbookLifecycleService == null)
                    {
                        _owner.QuitApplicationCore();
                    }
                }
                else if (completionAction == KernelHomeSessionCompletionAction.DismissPreparedHomeDisplayState)
                {
                    _owner._displayService.DismissPreparedHomeDisplayStateCore("CloseHomeSession.CaseCreationSkipRestore");
                }
                else
                {
                    _owner._displayService.ReleaseHomeDisplayCore(true);
                }

                _owner._displayService.LogKernelFlickerTrace(
                    "source=KernelWorkbookService action=close-home-session-end entryPoint="
                    + (entryPoint ?? string.Empty)
                    + ", result=completed, workbook="
                    + _owner._displayService.FormatWorkbookDescriptor(workbook)
                    + ", activeState="
                    + _owner._displayService.FormatActiveExcelState());
                _owner._bindingService.ClearHomeWorkbookBinding("CloseHomeSession.Completed");
                _owner._logger.Info("CloseHomeSession completed. saveKernelWorkbook=" + saveKernelWorkbook.ToString());
            }
        }

        private sealed class PendingHomeSessionClose
        {
            internal PendingHomeSessionClose(
                string workbookKey,
                bool saveKernelWorkbook,
                KernelHomeSessionCompletionAction completionAction,
                string entryPoint,
                bool deferSessionCompletionUntilFormClosed,
                bool backendCloseCompleted)
            {
                WorkbookKey = workbookKey ?? string.Empty;
                SaveKernelWorkbook = saveKernelWorkbook;
                CompletionAction = completionAction;
                EntryPoint = entryPoint ?? string.Empty;
                DeferSessionCompletionUntilFormClosed = deferSessionCompletionUntilFormClosed;
                BackendCloseCompleted = backendCloseCompleted;
            }

            internal string WorkbookKey { get; }

            internal bool SaveKernelWorkbook { get; }

            internal KernelHomeSessionCompletionAction CompletionAction { get; }

            internal string EntryPoint { get; }

            internal bool DeferSessionCompletionUntilFormClosed { get; }

            internal bool BackendCloseCompleted { get; private set; }

            internal void MarkBackendCloseCompleted()
            {
                BackendCloseCompleted = true;
            }
        }

        internal sealed class TestHooks
        {
            internal Action QuitApplication { get; set; }

            internal Func<Excel.Workbook, bool> RequestManagedCloseFromHomeExit { get; set; }

            internal Action<Excel.Workbook> SaveAndCloseKernelWorkbook { get; set; }

            internal Action<Excel.Workbook> CloseKernelWorkbookWithoutLifecycle { get; set; }
        }
    }
}
