using System;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn
{
    internal sealed class AddInCompositionRoot
    {
        private readonly ThisAddIn _addIn;
        private readonly Excel.Application _application;
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Func<string, Excel.Workbook, Excel.Window, bool> _isTaskPaneRefreshSucceeded;
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly Action _showKernelHomeFromKernelCommand;
        private readonly Action<Excel.Range> _clearKernelSheetCommandCell;
        private readonly Action<object> _releaseComObject;
        private readonly Action<string> _showKernelHomePlaceholderWithExternalWorkbookSuppression;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<string, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Action _scheduleWordWarmup;
        private readonly int _pendingPaneRefreshMaxAttempts;
        private readonly string _kernelSheetCommandSheetCodeName;
        private readonly string _kernelSheetCommandCellAddress;

        internal AddInCompositionRoot(
            ThisAddIn addIn,
            Excel.Application application,
            Logger logger,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Func<string, Excel.Workbook, Excel.Window, bool> isTaskPaneRefreshSucceeded,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount,
            Action showKernelHomeFromKernelCommand,
            Action<Excel.Range> clearKernelSheetCommandCell,
            Action<object> releaseComObject,
            Action<string> showKernelHomePlaceholderWithExternalWorkbookSuppression,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPane,
            Action scheduleWordWarmup,
            int pendingPaneRefreshMaxAttempts,
            string kernelSheetCommandSheetCodeName,
            string kernelSheetCommandCellAddress)
        {
            _addIn = addIn;
            _application = application;
            _logger = logger;
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow;
            _isTaskPaneRefreshSucceeded = isTaskPaneRefreshSucceeded;
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
            _showKernelHomeFromKernelCommand = showKernelHomeFromKernelCommand;
            _clearKernelSheetCommandCell = clearKernelSheetCommandCell;
            _releaseComObject = releaseComObject;
            _showKernelHomePlaceholderWithExternalWorkbookSuppression = showKernelHomePlaceholderWithExternalWorkbookSuppression;
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPane = refreshTaskPane;
            _scheduleWordWarmup = scheduleWordWarmup;
            _pendingPaneRefreshMaxAttempts = pendingPaneRefreshMaxAttempts;
            _kernelSheetCommandSheetCodeName = kernelSheetCommandSheetCodeName;
            _kernelSheetCommandCellAddress = kernelSheetCommandCellAddress;
        }

        internal ExcelInteropService ExcelInteropService { get; private set; }

        internal ExcelWindowRecoveryService ExcelWindowRecoveryService { get; private set; }

        internal WorkbookRoleResolver WorkbookRoleResolver { get; private set; }

        internal NavigationService NavigationService { get; private set; }

        internal WorkbookSessionService WorkbookSessionService { get; private set; }

        internal TransientPaneSuppressionService TransientPaneSuppressionService { get; private set; }

        internal CaseContextFactory CaseContextFactory { get; private set; }

        internal DocumentCommandService DocumentCommandService { get; private set; }

        internal DocumentNamePromptService DocumentNamePromptService { get; private set; }

        internal DocumentExecutionModeService DocumentExecutionModeService { get; private set; }

        internal WordInteropService WordInteropService { get; private set; }

        internal KernelWorkbookService KernelWorkbookService { get; private set; }

        internal KernelWorkbookLifecycleService KernelWorkbookLifecycleService { get; private set; }

        internal KernelCaseCreationCommandService KernelCaseCreationCommandService { get; private set; }

        internal KernelUserDataReflectionService KernelUserDataReflectionService { get; private set; }

        internal KernelCommandService KernelCommandService { get; private set; }

        internal CaseWorkbookLifecycleService CaseWorkbookLifecycleService { get; private set; }

        internal WorkbookClipboardPreservationService WorkbookClipboardPreservationService { get; private set; }

        internal AccountingSheetControlService AccountingSheetControlService { get; private set; }

        internal AccountingWorkbookLifecycleService AccountingWorkbookLifecycleService { get; private set; }

        internal KernelCaseInteractionState KernelCaseInteractionState { get; private set; }

        internal TaskPaneManager TaskPaneManager { get; private set; }

        internal WorkbookRibbonCommandService WorkbookRibbonCommandService { get; private set; }

        internal WorkbookCaseTaskPaneRefreshCommandService WorkbookCaseTaskPaneRefreshCommandService { get; private set; }

        internal WorkbookResetCommandService WorkbookResetCommandService { get; private set; }

        internal KernelSheetCommandTriggerService KernelSheetCommandTriggerService { get; private set; }

        internal KernelWorkbookAvailabilityService KernelWorkbookAvailabilityService { get; private set; }

        internal WorkbookEventCoordinator WorkbookEventCoordinator { get; private set; }

        internal KernelHomeCoordinator KernelHomeCoordinator { get; private set; }

        internal KernelHomeCasePaneSuppressionCoordinator KernelHomeCasePaneSuppressionCoordinator { get; private set; }

        internal ExternalWorkbookDetectionService ExternalWorkbookDetectionService { get; private set; }

        internal WindowActivatePaneHandlingService WindowActivatePaneHandlingService { get; private set; }

        internal TaskPaneRefreshOrchestrationService TaskPaneRefreshOrchestrationService { get; private set; }

        internal void Compose()
        {
            var pathCompatibilityService = new PathCompatibilityService();
            KernelCaseInteractionState = new KernelCaseInteractionState(_logger);
            ExcelInteropService = new ExcelInteropService(_application, _logger, pathCompatibilityService);
            var caseTaskPaneViewStateBuilder = new CaseTaskPaneViewStateBuilder();
            ExcelWindowRecoveryService = new ExcelWindowRecoveryService(_application, ExcelInteropService, _logger);
            var caseListFieldDefinitionRepository = new CaseListFieldDefinitionRepository(ExcelInteropService);
            var caseListHeaderRepository = new CaseListHeaderRepository(ExcelInteropService);
            var caseListMappingRepository = new CaseListMappingRepository(ExcelInteropService);
            var kernelWorkbookResolverService = new KernelWorkbookResolverService(_application, ExcelInteropService, pathCompatibilityService);
            var caseDataSnapshotFactory = new CaseDataSnapshotFactory(ExcelInteropService, kernelWorkbookResolverService, caseListFieldDefinitionRepository, _logger);
            TransientPaneSuppressionService = new TransientPaneSuppressionService(ExcelInteropService, pathCompatibilityService, _logger);
            WorkbookRoleResolver = new WorkbookRoleResolver(ExcelInteropService, pathCompatibilityService);
            WorkbookClipboardPreservationService = new WorkbookClipboardPreservationService(WorkbookRoleResolver, _logger);
            NavigationService = new NavigationService(ExcelInteropService, WorkbookRoleResolver, _logger);
            WorkbookSessionService = new WorkbookSessionService(NavigationService, TransientPaneSuppressionService, _logger);
            CaseContextFactory = new CaseContextFactory(ExcelInteropService, caseDataSnapshotFactory, _logger);
            KernelWorkbookService = new KernelWorkbookService(_application, ExcelInteropService, ExcelWindowRecoveryService, KernelCaseInteractionState, _logger);
            var userErrorService = new UserErrorService(_logger);
            KernelWorkbookLifecycleService = new KernelWorkbookLifecycleService(KernelWorkbookService, _application, ExcelInteropService, pathCompatibilityService, _logger);
            KernelWorkbookService.SetLifecycleService(KernelWorkbookLifecycleService);
            CaseWorkbookLifecycleService = new CaseWorkbookLifecycleService(WorkbookRoleResolver, _application, ExcelInteropService, pathCompatibilityService, TransientPaneSuppressionService, _logger);
            var folderWindowService = new FolderWindowService(pathCompatibilityService, _logger);
            var taskPaneSnapshotBuilderService = new TaskPaneSnapshotBuilderService(_application, ExcelInteropService, pathCompatibilityService, _logger);
            var kernelCasePathService = new KernelCasePathService(pathCompatibilityService);
            var taskPaneSnapshotCacheService = new TaskPaneSnapshotCacheService(ExcelInteropService, _logger);
            var masterTemplateCatalogService = new MasterTemplateCatalogService(_application, ExcelInteropService, pathCompatibilityService, _logger);
            var documentTemplateResolver = new DocumentTemplateResolver(ExcelInteropService, pathCompatibilityService, taskPaneSnapshotCacheService, masterTemplateCatalogService, _logger);
            var documentOutputService = new DocumentOutputService(ExcelInteropService, pathCompatibilityService, _logger);
            var excelValidationService = new ExcelValidationService(_logger);
            var accountingTemplateResolver = new AccountingTemplateResolver(ExcelInteropService, pathCompatibilityService, _logger);
            var accountingWorkbookService = new AccountingWorkbookService(_application, excelValidationService, _logger);
            AccountingSheetControlService = new AccountingSheetControlService(WorkbookRoleResolver, accountingWorkbookService, _logger);
            var accountingSetNamingService = new AccountingSetNamingService(documentOutputService, pathCompatibilityService);
            var accountingPaymentHistoryImportService = new AccountingPaymentHistoryImportService(accountingWorkbookService, userErrorService, _logger);
            var accountingSheetCommandService = new AccountingSheetCommandService(accountingWorkbookService, _logger);
            var accountingSaveAsService = new AccountingSaveAsService(ExcelInteropService, accountingWorkbookService, documentOutputService, pathCompatibilityService, userErrorService, _logger);
            var accountingInstallmentScheduleCommandService = new AccountingInstallmentScheduleCommandService(accountingWorkbookService, userErrorService, _logger);
            var accountingPaymentHistoryCommandService = new AccountingPaymentHistoryCommandService(accountingWorkbookService, userErrorService, _logger);
            var accountingFormHelperService = new AccountingFormHelperService(accountingWorkbookService, accountingInstallmentScheduleCommandService, accountingPaymentHistoryCommandService, accountingSaveAsService, userErrorService, _logger);
            AccountingWorkbookLifecycleService = new AccountingWorkbookLifecycleService(WorkbookRoleResolver, accountingWorkbookService, accountingFormHelperService, accountingPaymentHistoryImportService, _logger);
            var accountingInternalCommandService = new AccountingInternalCommandService(NavigationService, accountingPaymentHistoryImportService, accountingFormHelperService, accountingSaveAsService, _logger);
            DocumentExecutionModeService = new DocumentExecutionModeService(_logger, ExcelInteropService);
            var documentExecutionEligibilityService = new DocumentExecutionEligibilityService(ExcelInteropService, documentTemplateResolver, CaseContextFactory, documentOutputService, _logger);
            var documentExecutionPolicyService = new DocumentExecutionPolicyService(_logger, ExcelInteropService);
            var documentEligibilityDiagnosticsService = new DocumentEligibilityDiagnosticsService(DocumentExecutionModeService, documentExecutionEligibilityService, documentExecutionPolicyService, _logger);
            var documentMasterCatalogDiagnosticsService = new DocumentMasterCatalogDiagnosticsService(masterTemplateCatalogService, documentExecutionEligibilityService, documentExecutionPolicyService, DocumentExecutionModeService, _logger);
            var mergeDataBuilder = new MergeDataBuilder();
            var documentMergeService = new DocumentMergeService(_logger);
            WordInteropService = new WordInteropService(pathCompatibilityService, _logger);
            var localWorkCopyService = new LocalWorkCopyService(pathCompatibilityService, WordInteropService, _logger);
            var documentSaveService = new DocumentSaveService(documentOutputService, pathCompatibilityService, localWorkCopyService, WordInteropService, _logger);
            var documentCreateService = new DocumentCreateService(
                ExcelInteropService,
                CaseContextFactory,
                documentOutputService,
                mergeDataBuilder,
                documentMergeService,
                documentSaveService,
                WordInteropService,
                _logger);
            var accountingSetCreateService = new AccountingSetCreateService(
                ExcelInteropService,
                CaseContextFactory,
                documentOutputService,
                accountingSetNamingService,
                accountingTemplateResolver,
                accountingWorkbookService,
                pathCompatibilityService,
                TransientPaneSuppressionService,
                _logger);
            var accountingSetKernelSyncService = new AccountingSetKernelSyncService(
                ExcelInteropService,
                accountingTemplateResolver,
                accountingWorkbookService,
                pathCompatibilityService,
                _logger);
            var accountingSetCommandService = new AccountingSetCommandService(
                WorkbookRoleResolver,
                accountingSetCreateService,
                accountingSetKernelSyncService,
                _logger);
            var caseListRegistrationService = new CaseListRegistrationService(
                _application,
                ExcelInteropService,
                kernelWorkbookResolverService,
                caseDataSnapshotFactory,
                caseListFieldDefinitionRepository,
                caseListHeaderRepository,
                caseListMappingRepository,
                accountingWorkbookService,
                _logger);
            var screenUpdatingExecutionBridge = new ThisAddInScreenUpdatingExecutionBridge(_addIn);
            var taskPaneRefreshSuppressionBridge = new ThisAddInTaskPaneRefreshSuppressionBridge(_addIn);
            var activeTaskPaneRefreshBridge = new ThisAddInActiveTaskPaneRefreshBridge(_addIn);
            DocumentCommandService = new DocumentCommandService(_addIn, screenUpdatingExecutionBridge, taskPaneRefreshSuppressionBridge, activeTaskPaneRefreshBridge, DocumentExecutionModeService, documentExecutionEligibilityService, documentExecutionPolicyService, documentCreateService, accountingSetCommandService, caseListRegistrationService, CaseContextFactory, ExcelInteropService, _logger);
            DocumentNamePromptService = new DocumentNamePromptService(ExcelInteropService, taskPaneSnapshotCacheService, _logger);
            WorkbookRibbonCommandService = new WorkbookRibbonCommandService(ExcelInteropService, pathCompatibilityService, _logger);
            WorkbookCaseTaskPaneRefreshCommandService = new WorkbookCaseTaskPaneRefreshCommandService(
                WorkbookRoleResolver,
                ExcelInteropService,
                _resolveWorkbookPaneWindow,
                _isTaskPaneRefreshSucceeded);
            var workbookResetDefinitionRepository = new WorkbookResetDefinitionRepository();
            WorkbookResetCommandService = new WorkbookResetCommandService(
                ExcelInteropService,
                WorkbookRoleResolver,
                workbookResetDefinitionRepository,
                KernelWorkbookLifecycleService,
                _logger);
            var caseTemplateSnapshotService = new CaseTemplateSnapshotService(ExcelInteropService);
            var caseWorkbookInitializer = new CaseWorkbookInitializer(ExcelInteropService, caseTemplateSnapshotService, caseListFieldDefinitionRepository);
            var caseWorkbookOpenStrategy = new CaseWorkbookOpenStrategy(_application, WorkbookRoleResolver, _logger);
            var createdCaseOpenPromptService = new CreatedCaseOpenPromptService(_logger);
            var kernelCasePresentationService = new KernelCasePresentationService(_application, caseWorkbookOpenStrategy, ExcelInteropService, ExcelWindowRecoveryService, kernelWorkbookResolverService, caseListFieldDefinitionRepository, folderWindowService, TransientPaneSuppressionService, _logger);
            var kernelCaseCreationService = new KernelCaseCreationService(KernelWorkbookService, kernelCasePathService, caseWorkbookInitializer, caseWorkbookOpenStrategy, folderWindowService, TransientPaneSuppressionService, ExcelInteropService, _logger);
            KernelCaseCreationCommandService = new KernelCaseCreationCommandService(KernelWorkbookService, kernelCaseCreationService, kernelCasePathService, kernelCasePresentationService, createdCaseOpenPromptService, ExcelInteropService, _logger);
            KernelUserDataReflectionService = new KernelUserDataReflectionService(
                KernelWorkbookService,
                ExcelInteropService,
                accountingTemplateResolver,
                accountingWorkbookService,
                pathCompatibilityService,
                new UserDataBaseMappingRepository(ExcelInteropService),
                _logger);
            var kernelTemplateSyncService = new KernelTemplateSyncService(
                _application,
                KernelWorkbookService,
                ExcelInteropService,
                accountingWorkbookService,
                pathCompatibilityService,
                masterTemplateCatalogService,
                CaseWorkbookLifecycleService,
                _logger);
            KernelCommandService = new KernelCommandService(KernelWorkbookService, KernelUserDataReflectionService, kernelTemplateSyncService, _showKernelHomeFromKernelCommand, _logger);
            KernelSheetCommandTriggerService = new KernelSheetCommandTriggerService(
                KernelCommandService,
                KernelWorkbookService,
                ExcelInteropService,
                _application,
                _kernelSheetCommandSheetCodeName,
                _kernelSheetCommandCellAddress,
                _clearKernelSheetCommandCell,
                _releaseComObject,
                _logger);
            ExternalWorkbookDetectionService = new ExternalWorkbookDetectionService(
                WorkbookRoleResolver,
                KernelCaseInteractionState,
                KernelWorkbookService,
                TransientPaneSuppressionService,
                ExcelInteropService,
                _logger);
            WorkbookEventCoordinator = new WorkbookEventCoordinator(_addIn);
            KernelHomeCasePaneSuppressionCoordinator = new KernelHomeCasePaneSuppressionCoordinator(
                WorkbookRoleResolver,
                ExcelInteropService,
                _logger);
            KernelHomeCoordinator = new KernelHomeCoordinator(_addIn, KernelHomeCasePaneSuppressionCoordinator);
            KernelWorkbookAvailabilityService = new KernelWorkbookAvailabilityService(
                KernelWorkbookService,
                ExcelInteropService,
                KernelHomeCoordinator,
                _showKernelHomePlaceholderWithExternalWorkbookSuppression,
                _logger);
            var taskPaneDisplayRetryCoordinator = new TaskPaneDisplayRetryCoordinator(_pendingPaneRefreshMaxAttempts);
            var workbookTaskPaneDisplayAttemptCoordinator = new WorkbookTaskPaneDisplayAttemptCoordinator();
            TaskPaneManager = new TaskPaneManager(_addIn, ExcelInteropService, taskPaneSnapshotBuilderService, DocumentCommandService, documentEligibilityDiagnosticsService, documentMasterCatalogDiagnosticsService, DocumentNamePromptService, KernelCommandService, accountingSheetCommandService, caseTaskPaneViewStateBuilder, accountingInternalCommandService, KernelCaseInteractionState, userErrorService, _logger);
            WindowActivatePaneHandlingService = new WindowActivatePaneHandlingService(
                _handleExternalWorkbookDetected,
                _shouldSuppressCasePaneRefresh,
                TaskPaneManager,
                _refreshTaskPane,
                ExcelInteropService,
                _logger);
            var taskPaneRefreshCoordinator = new TaskPaneRefreshCoordinator(
                WorkbookSessionService,
                TaskPaneManager,
                ExcelWindowRecoveryService,
                _logger,
                _resolveWorkbookPaneWindow,
                _scheduleWordWarmup);
            TaskPaneRefreshOrchestrationService = new TaskPaneRefreshOrchestrationService(
                ExcelInteropService,
                WorkbookSessionService,
                _logger,
                taskPaneDisplayRetryCoordinator,
                workbookTaskPaneDisplayAttemptCoordinator,
                taskPaneRefreshCoordinator,
                _getKernelHomeForm,
                _getTaskPaneRefreshSuppressionCount);
        }
    }
}
