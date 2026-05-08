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
        private readonly Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> _tryRefreshTaskPane;
        private readonly Func<string, Excel.Workbook, Excel.Window, bool> _isTaskPaneRefreshSucceeded;
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly Action _showKernelHomeFromKernelCommand;
        private readonly Action<Excel.Range> _clearKernelSheetCommandCell;
        private readonly Action<object> _releaseComObject;
        private readonly Action<string> _showKernelHomePlaceholderWithExternalWorkbookSuppression;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<string, Excel.Workbook, Excel.Window> _refreshTaskPaneByReason;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Action _scheduleWordWarmup;
        private readonly int _pendingPaneRefreshMaxAttempts;
        private readonly string _kernelSheetCommandSheetCodeName;
        private readonly string _kernelSheetCommandCellAddress;

        internal AddInCompositionRoot(
            ThisAddIn addIn,
            Excel.Application application,
            Logger logger,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<string, Excel.Workbook, Excel.Window, bool> isTaskPaneRefreshSucceeded,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount,
            Action showKernelHomeFromKernelCommand,
            Action<Excel.Range> clearKernelSheetCommandCell,
            Action<object> releaseComObject,
            Action<string> showKernelHomePlaceholderWithExternalWorkbookSuppression,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<string, Excel.Workbook, Excel.Window> refreshTaskPaneByReason,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane,
            Action scheduleWordWarmup,
            int pendingPaneRefreshMaxAttempts,
            string kernelSheetCommandSheetCodeName,
            string kernelSheetCommandCellAddress)
        {
            _addIn = addIn;
            _application = application;
            _logger = logger;
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow;
            _tryRefreshTaskPane = tryRefreshTaskPane;
            _isTaskPaneRefreshSucceeded = isTaskPaneRefreshSucceeded;
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
            _showKernelHomeFromKernelCommand = showKernelHomeFromKernelCommand;
            _clearKernelSheetCommandCell = clearKernelSheetCommandCell;
            _releaseComObject = releaseComObject;
            _showKernelHomePlaceholderWithExternalWorkbookSuppression = showKernelHomePlaceholderWithExternalWorkbookSuppression;
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPaneByReason = refreshTaskPaneByReason;
            _refreshTaskPane = refreshTaskPane;
            _scheduleWordWarmup = scheduleWordWarmup;
            _pendingPaneRefreshMaxAttempts = pendingPaneRefreshMaxAttempts;
            _kernelSheetCommandSheetCodeName = kernelSheetCommandSheetCodeName;
            _kernelSheetCommandCellAddress = kernelSheetCommandCellAddress;
        }

        internal ExcelInteropService ExcelInteropService { get; private set; }

        internal WorkbookRoleResolver WorkbookRoleResolver { get; private set; }

        internal CaseWorkbookOpenStrategy CaseWorkbookOpenStrategy { get; private set; }

        internal DocumentExecutionModeService DocumentExecutionModeService { get; private set; }

        internal WordInteropService WordInteropService { get; private set; }

        internal KernelWorkbookService KernelWorkbookService { get; private set; }

        internal KernelWorkbookLifecycleService KernelWorkbookLifecycleService { get; private set; }

        internal KernelCaseCreationCommandService KernelCaseCreationCommandService { get; private set; }

        internal KernelUserDataReflectionService KernelUserDataReflectionService { get; private set; }

        internal KernelCaseInteractionState KernelCaseInteractionState { get; private set; }

        internal TaskPaneManager TaskPaneManager { get; private set; }

        internal WorkbookRibbonCommandService WorkbookRibbonCommandService { get; private set; }

        internal WorkbookCaseTaskPaneRefreshCommandService WorkbookCaseTaskPaneRefreshCommandService { get; private set; }

        internal WorkbookResetCommandService WorkbookResetCommandService { get; private set; }

        internal KernelWorkbookAvailabilityService KernelWorkbookAvailabilityService { get; private set; }

        internal WorkbookEventCoordinator WorkbookEventCoordinator { get; private set; }

        internal SheetEventCoordinator SheetEventCoordinator { get; private set; }

        internal WorkbookLifecycleCoordinator WorkbookLifecycleCoordinator { get; private set; }

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
            var workbookWindowVisibilityService = new WorkbookWindowVisibilityService(ExcelInteropService, _logger);
            var excelWindowRecoveryService = new ExcelWindowRecoveryService(_application, ExcelInteropService, _logger);
            var caseListFieldDefinitionRepository = new CaseListFieldDefinitionRepository(ExcelInteropService);
            var caseListHeaderRepository = new CaseListHeaderRepository(ExcelInteropService);
            var caseListMappingRepository = new CaseListMappingRepository(ExcelInteropService);
            var kernelWorkbookResolverService = new KernelWorkbookResolverService(_application, ExcelInteropService, pathCompatibilityService);
            var caseDataSnapshotFactory = new CaseDataSnapshotFactory(ExcelInteropService, kernelWorkbookResolverService, caseListFieldDefinitionRepository, _logger);
            var transientPaneSuppressionService = new TransientPaneSuppressionService(ExcelInteropService, pathCompatibilityService, _logger);
            WorkbookRoleResolver = new WorkbookRoleResolver(ExcelInteropService, pathCompatibilityService);
            var workbookClipboardPreservationService = new WorkbookClipboardPreservationService(WorkbookRoleResolver, _logger);
            var navigationService = new NavigationService(ExcelInteropService, WorkbookRoleResolver, _logger);
            var workbookSessionService = new WorkbookSessionService(navigationService, transientPaneSuppressionService, _logger);
            var caseContextFactory = new CaseContextFactory(ExcelInteropService, caseDataSnapshotFactory, _logger);
            // Dependency memo:
            // - Stable intra-area wiring is extracted into small compositions.
            // - Cross-area coordination and creation order stay in this root intentionally.
            // Kernel workbook boundary: workbook access and lifecycle only.
            var kernelWorkbookCoreComposition = new AddInKernelWorkbookCoreCompositionFactory(_application, _logger)
                .Compose(
                    pathCompatibilityService,
                    ExcelInteropService,
                    caseListFieldDefinitionRepository,
                    caseListHeaderRepository,
                    caseListMappingRepository,
                    excelWindowRecoveryService,
                    KernelCaseInteractionState);
            KernelWorkbookService = kernelWorkbookCoreComposition.KernelWorkbookService;
            KernelWorkbookLifecycleService = kernelWorkbookCoreComposition.KernelWorkbookLifecycleService;
            var userErrorService = new UserErrorService(_logger);
            var folderWindowService = new FolderWindowService(pathCompatibilityService, _logger);
            var kernelTemplateFolderPathResolver = new KernelTemplateFolderPathResolver(ExcelInteropService, pathCompatibilityService);
            var kernelTemplateFolderOpenService = new KernelTemplateFolderOpenService(
                kernelTemplateFolderPathResolver,
                pathCompatibilityService,
                folderWindowService,
                _logger);
            var managedCloseState = new ManagedCloseState();
            var caseFolderOpenService = new CaseFolderOpenService(ExcelInteropService, pathCompatibilityService, folderWindowService);
            var caseClosePromptService = new CaseClosePromptService(ExcelInteropService);
            var kernelNameRuleReader = new KernelNameRuleReader(ExcelInteropService, pathCompatibilityService, _logger);
            var postCloseFollowUpScheduler = new PostCloseFollowUpScheduler(_application, ExcelInteropService, _logger);
            var caseWorkbookLifecycleService = new CaseWorkbookLifecycleService(
                WorkbookRoleResolver,
                _application,
                ExcelInteropService,
                transientPaneSuppressionService,
                managedCloseState,
                caseClosePromptService,
                caseFolderOpenService,
                kernelNameRuleReader,
                postCloseFollowUpScheduler,
                _logger);
            var kernelCasePathService = new KernelCasePathService(pathCompatibilityService);
            var taskPaneSnapshotCacheService = new TaskPaneSnapshotCacheService(ExcelInteropService, _logger);
            var masterTemplateSheetReader = new MasterTemplateSheetReaderAdapter();
            var masterWorkbookReadAccessService = new MasterWorkbookReadAccessService(_application, ExcelInteropService, pathCompatibilityService);
            var masterTemplateCatalogService = new MasterTemplateCatalogService(ExcelInteropService, masterWorkbookReadAccessService, masterTemplateSheetReader, _logger);
            var wordTemplateContentControlInspectionService = new WordTemplateContentControlInspectionService();
            var wordTemplateRegistrationValidationService = new WordTemplateRegistrationValidationService(wordTemplateContentControlInspectionService, _logger);
            var kernelTemplateSyncPreflightService = new KernelTemplateSyncPreflightService(pathCompatibilityService, wordTemplateRegistrationValidationService);
            var kernelTemplateSyncPreparationService = new KernelTemplateSyncPreparationService(ExcelInteropService, pathCompatibilityService, caseListFieldDefinitionRepository, kernelTemplateSyncPreflightService);
            var documentOutputService = new DocumentOutputService(ExcelInteropService, pathCompatibilityService, _logger);
            var excelValidationService = new ExcelValidationService(_logger);
            var accountingTemplateResolver = new AccountingTemplateResolver(ExcelInteropService, pathCompatibilityService, _logger);
            var accountingWorkbookService = new AccountingWorkbookService(_application, excelValidationService, _logger);
            var accountingSheetControlService = new AccountingSheetControlService(WorkbookRoleResolver, accountingWorkbookService, _logger);
            var accountingSetNamingService = new AccountingSetNamingService(documentOutputService, pathCompatibilityService);
            var accountingPaymentHistoryImportService = new AccountingPaymentHistoryImportService(accountingWorkbookService, userErrorService, _logger);
            var accountingSheetCommandService = new AccountingSheetCommandService(accountingWorkbookService, _logger);
            var accountingSaveAsService = new AccountingSaveAsService(ExcelInteropService, accountingWorkbookService, documentOutputService, pathCompatibilityService, userErrorService, _logger);
            var accountingInstallmentScheduleCommandService = new AccountingInstallmentScheduleCommandService(accountingWorkbookService, userErrorService, _logger);
            var accountingPaymentHistoryCommandService = new AccountingPaymentHistoryCommandService(accountingWorkbookService, userErrorService, _logger);
            var accountingFormHelperService = new AccountingFormHelperService(accountingWorkbookService, accountingInstallmentScheduleCommandService, accountingPaymentHistoryCommandService, accountingSaveAsService, userErrorService, _logger);
            var accountingWorkbookLifecycleService = new AccountingWorkbookLifecycleService(WorkbookRoleResolver, accountingWorkbookService, accountingFormHelperService, accountingPaymentHistoryImportService, _logger);
            var accountingInternalCommandService = new AccountingInternalCommandService(navigationService, accountingPaymentHistoryImportService, accountingFormHelperService, accountingSaveAsService, _logger);
            var accountingSetPresentationWaitService = new AccountingSetPresentationWaitService(_logger);
            var accountingSetReadyShowBridge = new ThisAddInAccountingSetReadyShowBridge(_addIn);
            var accountingSetCreateService = new AccountingSetCreateService(
                ExcelInteropService,
                caseContextFactory,
                documentOutputService,
                accountingSetNamingService,
                accountingTemplateResolver,
                accountingWorkbookService,
                pathCompatibilityService,
                transientPaneSuppressionService,
                accountingSetPresentationWaitService,
                accountingSetReadyShowBridge,
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
                ExcelInteropService,
                kernelWorkbookResolverService,
                caseDataSnapshotFactory,
                caseListFieldDefinitionRepository,
                caseListHeaderRepository,
                caseListMappingRepository,
                accountingWorkbookService,
                taskPaneSnapshotCacheService,
                _logger);
            // Document boundary: bundle Word/document execution services and diagnostics.
            var documentComposition = new AddInDocumentCompositionFactory(_addIn, _logger)
                .Compose(
                    pathCompatibilityService,
                    ExcelInteropService,
                    caseContextFactory,
                    taskPaneSnapshotCacheService,
                    masterTemplateCatalogService,
                    documentOutputService,
                    accountingSetCommandService,
                    caseListRegistrationService);
            var documentCommandService = documentComposition.DocumentCommandService;
            var documentNamePromptService = documentComposition.DocumentNamePromptService;
            DocumentExecutionModeService = documentComposition.DocumentExecutionModeService;
            WordInteropService = documentComposition.WordInteropService;
            WorkbookRibbonCommandService = new WorkbookRibbonCommandService(ExcelInteropService, pathCompatibilityService, _logger);
            var workbookResetDefinitionRepository = new WorkbookResetDefinitionRepository();
            WorkbookResetCommandService = new WorkbookResetCommandService(
                ExcelInteropService,
                WorkbookRoleResolver,
                workbookResetDefinitionRepository,
                KernelWorkbookLifecycleService,
                _logger);
            var caseTemplateSnapshotService = new CaseTemplateSnapshotService(ExcelInteropService);
            // Case creation stays here because it still spans kernel, case workbook, and UI-facing coordination.
            var caseWorkbookInitializer = new CaseWorkbookInitializer(ExcelInteropService, caseTemplateSnapshotService, caseListFieldDefinitionRepository);
            var caseWorkbookOpenStrategy = new CaseWorkbookOpenStrategy(_application, WorkbookRoleResolver, _logger);
            CaseWorkbookOpenStrategy = caseWorkbookOpenStrategy;
            var createdCasePresentationWaitService = new CreatedCasePresentationWaitService(_logger);
            var kernelUserDataRegistrationWaitService = new KernelUserDataRegistrationWaitService(_logger);
            var casePaneHostBridge = new ThisAddInCasePaneHostBridge(_addIn);
            var kernelCasePresentationService = new KernelCasePresentationService(_application, caseWorkbookOpenStrategy, ExcelInteropService, excelWindowRecoveryService, kernelWorkbookResolverService, caseListFieldDefinitionRepository, folderWindowService, createdCasePresentationWaitService, transientPaneSuppressionService, casePaneHostBridge, workbookWindowVisibilityService, _logger);
            var kernelCaseCreationService = new KernelCaseCreationService(KernelWorkbookService, kernelCasePathService, caseWorkbookInitializer, caseWorkbookOpenStrategy, transientPaneSuppressionService, caseWorkbookLifecycleService, ExcelInteropService, _logger);
            KernelCaseCreationCommandService = new KernelCaseCreationCommandService(KernelWorkbookService, kernelCaseCreationService, kernelCasePathService, kernelCasePresentationService, createdCasePresentationWaitService, caseWorkbookLifecycleService, ExcelInteropService, _logger);
            KernelUserDataReflectionService = new KernelUserDataReflectionService(
                KernelWorkbookService,
                ExcelInteropService,
                accountingTemplateResolver,
                accountingWorkbookService,
                pathCompatibilityService,
                new UserDataBaseMappingRepository(ExcelInteropService),
                _logger);
            var kernelUserDataRegistrationExecutionService = new KernelUserDataRegistrationExecutionService(
                KernelUserDataReflectionService,
                kernelUserDataRegistrationWaitService,
                _logger);
            var kernelTemplateSyncService = new KernelTemplateSyncService(
                _application,
                KernelWorkbookService,
                ExcelInteropService,
                accountingWorkbookService,
                pathCompatibilityService,
                kernelTemplateSyncPreparationService,
                masterTemplateCatalogService,
                caseWorkbookLifecycleService,
                _logger);
            var kernelCommandService = new KernelCommandService(
                KernelWorkbookService,
                KernelUserDataReflectionService,
                kernelUserDataRegistrationExecutionService,
                kernelTemplateSyncService,
                kernelTemplateFolderOpenService,
                _showKernelHomeFromKernelCommand,
                _logger);
            var kernelSheetCommandTriggerService = new KernelSheetCommandTriggerService(
                kernelCommandService,
                KernelWorkbookService,
                ExcelInteropService,
                _application,
                _kernelSheetCommandSheetCodeName,
                _kernelSheetCommandCellAddress,
                _clearKernelSheetCommandCell,
                _releaseComObject,
                _logger);
            // Kernel home remains here because it bridges workbook state and UI coordination.
            ExternalWorkbookDetectionService = new ExternalWorkbookDetectionService(
                WorkbookRoleResolver,
                KernelCaseInteractionState,
                KernelWorkbookService,
                transientPaneSuppressionService,
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
            // Task pane boundary: bundle pane construction and refresh orchestration.
            var taskPaneComposition = new AddInTaskPaneCompositionFactory(
                _addIn,
                _application,
                _logger,
                _resolveWorkbookPaneWindow,
                _tryRefreshTaskPane,
                _isTaskPaneRefreshSucceeded,
                _handleExternalWorkbookDetected,
                _shouldSuppressCasePaneRefresh,
                _refreshTaskPane,
                _scheduleWordWarmup,
                _getKernelHomeForm,
                _getTaskPaneRefreshSuppressionCount,
                casePaneHostBridge,
                _pendingPaneRefreshMaxAttempts)
                .Compose(
                    pathCompatibilityService,
                    ExcelInteropService,
                    masterWorkbookReadAccessService,
                    masterTemplateSheetReader,
                    WorkbookRoleResolver,
                    documentCommandService,
                    documentNamePromptService,
                    kernelCommandService,
                    accountingSheetCommandService,
                    accountingInternalCommandService,
                    KernelCaseInteractionState,
                    userErrorService,
                    workbookSessionService,
                    excelWindowRecoveryService,
                    workbookWindowVisibilityService);
            WorkbookCaseTaskPaneRefreshCommandService = taskPaneComposition.WorkbookCaseTaskPaneRefreshCommandService;
            TaskPaneManager = taskPaneComposition.TaskPaneManager;
            WindowActivatePaneHandlingService = taskPaneComposition.WindowActivatePaneHandlingService;
            TaskPaneRefreshOrchestrationService = taskPaneComposition.TaskPaneRefreshOrchestrationService;
            SheetEventCoordinator = new SheetEventCoordinator(
                _logger,
                kernelSheetCommandTriggerService,
                caseWorkbookLifecycleService,
                accountingWorkbookLifecycleService,
                accountingSheetControlService,
                _refreshTaskPaneByReason);
            WorkbookLifecycleCoordinator = new WorkbookLifecycleCoordinator(
                _logger,
                ExcelInteropService,
                KernelWorkbookLifecycleService,
                caseWorkbookLifecycleService,
                accountingWorkbookLifecycleService,
                accountingSheetControlService,
                workbookClipboardPreservationService,
                TaskPaneManager,
                KernelHomeCoordinator,
                _handleExternalWorkbookDetected,
                _refreshTaskPaneByReason,
                _shouldSuppressCasePaneRefresh,
                casePaneHostBridge);
        }
    }

    // Bundles document creation, execution diagnostics, and Word interop services.
    internal sealed class AddInDocumentCompositionFactory
    {
        private readonly ThisAddIn _addIn;
        private readonly Logger _logger;

        internal AddInDocumentCompositionFactory(ThisAddIn addIn, Logger logger)
        {
            _addIn = addIn;
            _logger = logger;
        }

        internal AddInDocumentComposition Compose(
            PathCompatibilityService pathCompatibilityService,
            ExcelInteropService excelInteropService,
            CaseContextFactory caseContextFactory,
            TaskPaneSnapshotCacheService taskPaneSnapshotCacheService,
            MasterTemplateCatalogService masterTemplateCatalogService,
            DocumentOutputService documentOutputService,
            AccountingSetCommandService accountingSetCommandService,
            CaseListRegistrationService caseListRegistrationService)
        {
            IMasterTemplateCatalogReader masterTemplateCatalogReader = masterTemplateCatalogService;
            var documentTemplateLookupService = new DocumentTemplateLookupService(
                taskPaneSnapshotCacheService,
                masterTemplateCatalogReader);
            IDocumentTemplateLookupReader documentTemplateLookupReader = documentTemplateLookupService;
            ICaseCacheDocumentTemplateReader caseCacheDocumentTemplateReader = documentTemplateLookupService;
            var documentTemplateResolver = new DocumentTemplateResolver(
                excelInteropService,
                pathCompatibilityService,
                documentTemplateLookupReader,
                _logger);
            var documentExecutionModeService = new DocumentExecutionModeService(_logger, excelInteropService);
            var documentExecutionEligibilityService = new DocumentExecutionEligibilityService(
                documentTemplateResolver,
                caseContextFactory,
                documentOutputService,
                _logger);
            var mergeDataBuilder = new MergeDataBuilder();
            var documentMergeService = new DocumentMergeService(_logger);
            var wordInteropService = new WordInteropService(pathCompatibilityService, _logger);
            var documentPresentationWaitService = new DocumentPresentationWaitService(_logger);
            var documentSaveService = new DocumentSaveService(
                documentOutputService,
                wordInteropService,
                _logger);
            var documentCreateHostBridge = new ThisAddInDocumentCreateHostBridge(_addIn);
            var documentCreateService = new DocumentCreateService(
                excelInteropService,
                caseContextFactory,
                documentOutputService,
                mergeDataBuilder,
                documentMergeService,
                documentSaveService,
                wordInteropService,
                documentPresentationWaitService,
                documentCreateHostBridge,
                _logger);
            var screenUpdatingExecutionBridge = new ThisAddInScreenUpdatingExecutionBridge(_addIn);
            var taskPaneRefreshSuppressionBridge = new ThisAddInTaskPaneRefreshSuppressionBridge(_addIn);
            var activeTaskPaneRefreshBridge = new ThisAddInActiveTaskPaneRefreshBridge(_addIn);
            var kernelSheetPaneRefreshBridge = new ThisAddInKernelSheetPaneRefreshBridge(_addIn);
            var documentCommandService = new DocumentCommandService(
                screenUpdatingExecutionBridge,
                taskPaneRefreshSuppressionBridge,
                activeTaskPaneRefreshBridge,
                kernelSheetPaneRefreshBridge,
                documentExecutionModeService,
                documentExecutionEligibilityService,
                documentCreateService,
                accountingSetCommandService,
                caseListRegistrationService,
                caseContextFactory,
                excelInteropService,
                _logger);
            var documentNamePromptService = new DocumentNamePromptService(
                excelInteropService,
                caseCacheDocumentTemplateReader,
                _logger);
            return new AddInDocumentComposition(
                documentCommandService,
                documentNamePromptService,
                documentExecutionModeService,
                wordInteropService);
        }
    }

    // Carries the document-related services composed for the root.
    internal sealed class AddInDocumentComposition
    {
        internal AddInDocumentComposition(
            DocumentCommandService documentCommandService,
            DocumentNamePromptService documentNamePromptService,
            DocumentExecutionModeService documentExecutionModeService,
            WordInteropService wordInteropService)
        {
            DocumentCommandService = documentCommandService;
            DocumentNamePromptService = documentNamePromptService;
            DocumentExecutionModeService = documentExecutionModeService;
            WordInteropService = wordInteropService;
        }

        internal DocumentCommandService DocumentCommandService { get; private set; }

        internal DocumentNamePromptService DocumentNamePromptService { get; private set; }

        internal DocumentExecutionModeService DocumentExecutionModeService { get; private set; }

        internal WordInteropService WordInteropService { get; private set; }
    }

    // Bundles kernel workbook access and lifecycle services that must be initialized together.
    internal sealed class AddInKernelWorkbookCoreCompositionFactory
    {
        private readonly Excel.Application _application;
        private readonly Logger _logger;

        internal AddInKernelWorkbookCoreCompositionFactory(Excel.Application application, Logger logger)
        {
            _application = application;
            _logger = logger;
        }

        internal AddInKernelWorkbookCoreComposition Compose(
            PathCompatibilityService pathCompatibilityService,
            ExcelInteropService excelInteropService,
            CaseListFieldDefinitionRepository caseListFieldDefinitionRepository,
            CaseListHeaderRepository caseListHeaderRepository,
            CaseListMappingRepository caseListMappingRepository,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            KernelCaseInteractionState kernelCaseInteractionState)
        {
            var kernelWorkbookBindingService = new KernelWorkbookBindingService(
                _application,
                excelInteropService,
                pathCompatibilityService,
                _logger);
            var kernelWorkbookDisplayService = new KernelWorkbookDisplayService(
                _application,
                excelInteropService,
                excelWindowRecoveryService,
                kernelCaseInteractionState,
                _logger,
                kernelWorkbookBindingService);
            var kernelWorkbookCloseService = new KernelWorkbookCloseService(
                _application,
                kernelCaseInteractionState,
                _logger,
                kernelWorkbookBindingService,
                kernelWorkbookDisplayService);
            var kernelWorkbookService = new KernelWorkbookService(
                kernelWorkbookBindingService,
                kernelWorkbookDisplayService,
                kernelWorkbookCloseService);
            var kernelWorkbookLifecycleService = new KernelWorkbookLifecycleService(
                kernelWorkbookService,
                _application,
                excelInteropService,
                pathCompatibilityService,
                caseListFieldDefinitionRepository,
                caseListHeaderRepository,
                caseListMappingRepository,
                _logger);
            kernelWorkbookService.SetLifecycleService(kernelWorkbookLifecycleService);
            return new AddInKernelWorkbookCoreComposition(
                kernelWorkbookService,
                kernelWorkbookLifecycleService);
        }
    }

    // Carries kernel workbook core services back to the root.
    internal sealed class AddInKernelWorkbookCoreComposition
    {
        internal AddInKernelWorkbookCoreComposition(
            KernelWorkbookService kernelWorkbookService,
            KernelWorkbookLifecycleService kernelWorkbookLifecycleService)
        {
            KernelWorkbookService = kernelWorkbookService;
            KernelWorkbookLifecycleService = kernelWorkbookLifecycleService;
        }

        internal KernelWorkbookService KernelWorkbookService { get; private set; }

        internal KernelWorkbookLifecycleService KernelWorkbookLifecycleService { get; private set; }
    }

    // Bundles task pane construction and refresh orchestration services.
    internal sealed class AddInTaskPaneCompositionFactory
    {
        private readonly ThisAddIn _addIn;
        private readonly Excel.Application _application;
        private readonly Logger _logger;
        private readonly Func<Excel.Workbook, string, bool, Excel.Window> _resolveWorkbookPaneWindow;
        private readonly Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> _tryRefreshTaskPane;
        private readonly Func<string, Excel.Workbook, Excel.Window, bool> _isTaskPaneRefreshSucceeded;
        private readonly Action<Excel.Workbook, string> _handleExternalWorkbookDetected;
        private readonly Func<string, Excel.Workbook, bool> _shouldSuppressCasePaneRefresh;
        private readonly Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> _refreshTaskPane;
        private readonly Action _scheduleWordWarmup;
        private readonly Func<KernelHomeForm> _getKernelHomeForm;
        private readonly Func<int> _getTaskPaneRefreshSuppressionCount;
        private readonly ICasePaneHostBridge _casePaneHostBridge;
        private readonly int _pendingPaneRefreshMaxAttempts;

        internal AddInTaskPaneCompositionFactory(
            ThisAddIn addIn,
            Excel.Application application,
            Logger logger,
            Func<Excel.Workbook, string, bool, Excel.Window> resolveWorkbookPaneWindow,
            Func<string, Excel.Workbook, Excel.Window, TaskPaneRefreshAttemptResult> tryRefreshTaskPane,
            Func<string, Excel.Workbook, Excel.Window, bool> isTaskPaneRefreshSucceeded,
            Action<Excel.Workbook, string> handleExternalWorkbookDetected,
            Func<string, Excel.Workbook, bool> shouldSuppressCasePaneRefresh,
            Action<TaskPaneDisplayRequest, Excel.Workbook, Excel.Window> refreshTaskPane,
            Action scheduleWordWarmup,
            Func<KernelHomeForm> getKernelHomeForm,
            Func<int> getTaskPaneRefreshSuppressionCount,
            ICasePaneHostBridge casePaneHostBridge,
            int pendingPaneRefreshMaxAttempts)
        {
            _addIn = addIn;
            _application = application;
            _logger = logger;
            _resolveWorkbookPaneWindow = resolveWorkbookPaneWindow;
            _tryRefreshTaskPane = tryRefreshTaskPane;
            _isTaskPaneRefreshSucceeded = isTaskPaneRefreshSucceeded;
            _handleExternalWorkbookDetected = handleExternalWorkbookDetected;
            _shouldSuppressCasePaneRefresh = shouldSuppressCasePaneRefresh;
            _refreshTaskPane = refreshTaskPane;
            _scheduleWordWarmup = scheduleWordWarmup;
            _getKernelHomeForm = getKernelHomeForm;
            _getTaskPaneRefreshSuppressionCount = getTaskPaneRefreshSuppressionCount;
            _casePaneHostBridge = casePaneHostBridge ?? throw new ArgumentNullException(nameof(casePaneHostBridge));
            _pendingPaneRefreshMaxAttempts = pendingPaneRefreshMaxAttempts;
        }

        internal AddInTaskPaneComposition Compose(
            PathCompatibilityService pathCompatibilityService,
            ExcelInteropService excelInteropService,
            MasterWorkbookReadAccessService masterWorkbookReadAccessService,
            IMasterTemplateSheetReader masterTemplateSheetReader,
            WorkbookRoleResolver workbookRoleResolver,
            DocumentCommandService documentCommandService,
            DocumentNamePromptService documentNamePromptService,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            WorkbookSessionService workbookSessionService,
            ExcelWindowRecoveryService excelWindowRecoveryService,
            WorkbookWindowVisibilityService workbookWindowVisibilityService)
        {
            var caseTaskPaneViewStateBuilder = new CaseTaskPaneViewStateBuilder();
            var taskPaneSnapshotBuilderService = new TaskPaneSnapshotBuilderService(
                excelInteropService,
                pathCompatibilityService,
                masterWorkbookReadAccessService,
                masterTemplateSheetReader,
                _logger);
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader = taskPaneSnapshotBuilderService;
            var casePaneSnapshotRenderService = new CasePaneSnapshotRenderService(
                caseTaskPaneSnapshotReader,
                caseTaskPaneViewStateBuilder);
            var workbookCaseTaskPaneRefreshCommandService = new WorkbookCaseTaskPaneRefreshCommandService(
                workbookRoleResolver,
                excelInteropService,
                _resolveWorkbookPaneWindow,
                _isTaskPaneRefreshSucceeded);
            var taskPaneDisplayRetryCoordinator = new TaskPaneDisplayRetryCoordinator(_pendingPaneRefreshMaxAttempts);
            var workbookTaskPaneDisplayAttemptCoordinator = new WorkbookTaskPaneDisplayAttemptCoordinator();
            var workbookTaskPaneReadyShowAttemptWorker = new WorkbookTaskPaneReadyShowAttemptWorker(
                excelInteropService,
                _logger,
                taskPaneDisplayRetryCoordinator,
                workbookTaskPaneDisplayAttemptCoordinator,
                workbookWindowVisibilityService,
                (workbook, window) => _casePaneHostBridge.HasVisibleCasePaneForWorkbookWindow(workbook, window),
                _tryRefreshTaskPane,
                _resolveWorkbookPaneWindow);
            var taskPaneBusinessActionLauncher = new TaskPaneBusinessActionLauncher(
                documentCommandService,
                documentNamePromptService);
            var taskPaneManager = TaskPaneManagerRuntimeBootstrap.CreateAttached(
                _addIn,
                excelInteropService,
                taskPaneBusinessActionLauncher,
                kernelCommandService,
                accountingSheetCommandService,
                caseTaskPaneViewStateBuilder,
                casePaneSnapshotRenderService,
                accountingInternalCommandService,
                kernelCaseInteractionState,
                userErrorService,
                _logger);
            var windowActivatePanePredicateBridge = new ThisAddInWindowActivatePanePredicateBridge(_addIn);
            var windowActivatePaneHandlingService = new WindowActivatePaneHandlingService(
                windowActivatePanePredicateBridge,
                _handleExternalWorkbookDetected,
                _shouldSuppressCasePaneRefresh,
                _refreshTaskPane,
                _logger);
            var taskPaneRefreshCoordinator = new TaskPaneRefreshCoordinator(
                workbookSessionService,
                taskPaneManager,
                excelWindowRecoveryService,
                _logger,
                _resolveWorkbookPaneWindow,
                _scheduleWordWarmup,
                _casePaneHostBridge);
            var taskPaneRefreshOrchestrationService = new TaskPaneRefreshOrchestrationService(
                excelInteropService,
                workbookSessionService,
                _logger,
                taskPaneRefreshCoordinator,
                workbookTaskPaneReadyShowAttemptWorker,
                _getKernelHomeForm,
                _getTaskPaneRefreshSuppressionCount,
                _casePaneHostBridge);
            return new AddInTaskPaneComposition(
                workbookCaseTaskPaneRefreshCommandService,
                taskPaneManager,
                windowActivatePaneHandlingService,
                taskPaneRefreshOrchestrationService);
        }
    }

    // Carries task pane services back to the root.
    internal sealed class AddInTaskPaneComposition
    {
        internal AddInTaskPaneComposition(
            WorkbookCaseTaskPaneRefreshCommandService workbookCaseTaskPaneRefreshCommandService,
            TaskPaneManager taskPaneManager,
            WindowActivatePaneHandlingService windowActivatePaneHandlingService,
            TaskPaneRefreshOrchestrationService taskPaneRefreshOrchestrationService)
        {
            WorkbookCaseTaskPaneRefreshCommandService = workbookCaseTaskPaneRefreshCommandService;
            TaskPaneManager = taskPaneManager;
            WindowActivatePaneHandlingService = windowActivatePaneHandlingService;
            TaskPaneRefreshOrchestrationService = taskPaneRefreshOrchestrationService;
        }

        internal WorkbookCaseTaskPaneRefreshCommandService WorkbookCaseTaskPaneRefreshCommandService { get; private set; }

        internal TaskPaneManager TaskPaneManager { get; private set; }

        internal WindowActivatePaneHandlingService WindowActivatePaneHandlingService { get; private set; }

        internal TaskPaneRefreshOrchestrationService TaskPaneRefreshOrchestrationService { get; private set; }
    }
}
