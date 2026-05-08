using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // Manager attach payload: only collaborators that TaskPaneManager directly consumes at runtime live here.
    // Registration ownership stays on the compose/lifecycle side and is not routed back through manager attach.
    internal sealed class TaskPaneManagerRuntimeGraph
    {
        internal TaskPaneManagerRuntimeGraph(
            CasePaneCacheRefreshNotificationService casePaneCacheRefreshNotificationService,
            TaskPaneHostLifecycleService taskPaneHostLifecycleService,
            TaskPaneDisplayCoordinator taskPaneDisplayCoordinator,
            TaskPaneHostFlowService taskPaneHostFlowService)
        {
            CasePaneCacheRefreshNotificationService = casePaneCacheRefreshNotificationService ?? throw new ArgumentNullException(nameof(casePaneCacheRefreshNotificationService));
            TaskPaneHostLifecycleService = taskPaneHostLifecycleService ?? throw new ArgumentNullException(nameof(taskPaneHostLifecycleService));
            TaskPaneDisplayCoordinator = taskPaneDisplayCoordinator ?? throw new ArgumentNullException(nameof(taskPaneDisplayCoordinator));
            TaskPaneHostFlowService = taskPaneHostFlowService ?? throw new ArgumentNullException(nameof(taskPaneHostFlowService));
        }

        internal CasePaneCacheRefreshNotificationService CasePaneCacheRefreshNotificationService { get; }

        internal TaskPaneHostLifecycleService TaskPaneHostLifecycleService { get; }

        internal TaskPaneDisplayCoordinator TaskPaneDisplayCoordinator { get; }

        internal TaskPaneHostFlowService TaskPaneHostFlowService { get; }
    }

    // Immutable bootstrap entry input. This carries only the construction-time data needed to build an
    // unattached manager and derive a smaller graph-compose context; it never owns runtime state.
    internal sealed class TaskPaneManagerRuntimeEntryContext
    {
        private TaskPaneManagerRuntimeEntryContext(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks,
            bool usesFullRuntimeConstruction)
        {
            AddIn = addIn;
            ExcelInteropService = excelInteropService;
            TaskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher;
            KernelCommandService = kernelCommandService;
            AccountingSheetCommandService = accountingSheetCommandService;
            CaseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder;
            CasePaneSnapshotRenderService = casePaneSnapshotRenderService;
            AccountingInternalCommandService = accountingInternalCommandService;
            KernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            UserErrorService = userErrorService;
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            TestHooks = testHooks;
            UsesFullRuntimeConstruction = usesFullRuntimeConstruction;
        }

        internal static TaskPaneManagerRuntimeEntryContext CreateProductionFull(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger)
        {
            return CreateFullForTests(
                addIn,
                excelInteropService,
                taskPaneBusinessActionLauncher,
                kernelCommandService,
                accountingSheetCommandService,
                caseTaskPaneViewStateBuilder,
                casePaneSnapshotRenderService,
                accountingInternalCommandService,
                kernelCaseInteractionState,
                userErrorService,
                    logger,
                    testHooks: null);
        }

        internal static TaskPaneManagerRuntimeEntryContext CreateFullForTests(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks)
        {
            return new TaskPaneManagerRuntimeEntryContext(
                addIn ?? throw new ArgumentNullException(nameof(addIn)),
                excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService)),
                taskPaneBusinessActionLauncher ?? throw new ArgumentNullException(nameof(taskPaneBusinessActionLauncher)),
                kernelCommandService ?? throw new ArgumentNullException(nameof(kernelCommandService)),
                accountingSheetCommandService ?? throw new ArgumentNullException(nameof(accountingSheetCommandService)),
                caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder)),
                casePaneSnapshotRenderService ?? throw new ArgumentNullException(nameof(casePaneSnapshotRenderService)),
                accountingInternalCommandService ?? throw new ArgumentNullException(nameof(accountingInternalCommandService)),
                kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState)),
                userErrorService ?? throw new ArgumentNullException(nameof(userErrorService)),
                logger ?? throw new ArgumentNullException(nameof(logger)),
                testHooks,
                usesFullRuntimeConstruction: true);
        }

        internal static TaskPaneManagerRuntimeEntryContext CreateThinForTests(
            Logger logger,
            KernelCaseInteractionState kernelCaseInteractionState,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks)
        {
            return new TaskPaneManagerRuntimeEntryContext(
                addIn: null,
                excelInteropService: null,
                taskPaneBusinessActionLauncher: null,
                kernelCommandService: null,
                accountingSheetCommandService: null,
                caseTaskPaneViewStateBuilder: null,
                casePaneSnapshotRenderService: null,
                accountingInternalCommandService: null,
                kernelCaseInteractionState: kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState)),
                userErrorService: null,
                logger: logger ?? throw new ArgumentNullException(nameof(logger)),
                testHooks: testHooks,
                usesFullRuntimeConstruction: false);
        }

        internal TaskPaneManagerRuntimeGraphComposeContext CreateGraphComposeContext()
        {
            return new TaskPaneManagerRuntimeGraphComposeContext(
                AddIn,
                ExcelInteropService,
                TaskPaneBusinessActionLauncher,
                KernelCommandService,
                AccountingSheetCommandService,
                CaseTaskPaneViewStateBuilder,
                CasePaneSnapshotRenderService,
                AccountingInternalCommandService,
                KernelCaseInteractionState,
                UserErrorService,
                Logger,
                TestHooks);
        }

        internal ThisAddIn AddIn { get; }

        internal ExcelInteropService ExcelInteropService { get; }

        internal TaskPaneBusinessActionLauncher TaskPaneBusinessActionLauncher { get; }

        internal KernelCommandService KernelCommandService { get; }

        internal AccountingSheetCommandService AccountingSheetCommandService { get; }

        internal CaseTaskPaneViewStateBuilder CaseTaskPaneViewStateBuilder { get; }

        internal CasePaneSnapshotRenderService CasePaneSnapshotRenderService { get; }

        internal AccountingInternalCommandService AccountingInternalCommandService { get; }

        internal KernelCaseInteractionState KernelCaseInteractionState { get; }

        internal UserErrorService UserErrorService { get; }

        internal Logger Logger { get; }

        internal TaskPaneManager.TaskPaneManagerTestHooks TestHooks { get; }

        internal bool UsesFullRuntimeConstruction { get; }
    }

    // Passive graph-compose input. Intentionally excludes manager-construction-only dependencies such as the
    // snapshot reader so the graph factory does not become a second runtime entrypoint.
    internal sealed class TaskPaneManagerRuntimeGraphComposeContext
    {
        internal TaskPaneManagerRuntimeGraphComposeContext(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks)
        {
            AddIn = addIn;
            ExcelInteropService = excelInteropService;
            TaskPaneBusinessActionLauncher = taskPaneBusinessActionLauncher;
            KernelCommandService = kernelCommandService;
            AccountingSheetCommandService = accountingSheetCommandService;
            CaseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder;
            CasePaneSnapshotRenderService = casePaneSnapshotRenderService;
            AccountingInternalCommandService = accountingInternalCommandService;
            KernelCaseInteractionState = kernelCaseInteractionState ?? throw new ArgumentNullException(nameof(kernelCaseInteractionState));
            UserErrorService = userErrorService;
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            TestHooks = testHooks;
        }

        internal ThisAddIn AddIn { get; }

        internal ExcelInteropService ExcelInteropService { get; }

        internal TaskPaneBusinessActionLauncher TaskPaneBusinessActionLauncher { get; }

        internal KernelCommandService KernelCommandService { get; }

        internal AccountingSheetCommandService AccountingSheetCommandService { get; }

        internal CaseTaskPaneViewStateBuilder CaseTaskPaneViewStateBuilder { get; }

        internal CasePaneSnapshotRenderService CasePaneSnapshotRenderService { get; }

        internal AccountingInternalCommandService AccountingInternalCommandService { get; }

        internal KernelCaseInteractionState KernelCaseInteractionState { get; }

        internal UserErrorService UserErrorService { get; }

        internal Logger Logger { get; }

        internal TaskPaneManager.TaskPaneManagerTestHooks TestHooks { get; }
    }

    // Bootstrap-created compose surface over the manager-owned runtime state and render seam.
    // Graph assembly reads this explicit surface instead of treating TaskPaneManager itself as a dependency bag.
    internal sealed class TaskPaneManagerRuntimeGraphComposeSurface
    {
        internal TaskPaneManagerRuntimeGraphComposeSurface(
            Dictionary<string, TaskPaneHost> hostsByWindowKey,
            Func<WorkbookContext, string> formatContextDescriptor,
            Func<TaskPaneHost, string> formatHostDescriptor,
            Func<Excel.Workbook, string> formatWorkbookDescriptor,
            Func<TaskPaneHost, WorkbookContext, string, TaskPaneSnapshotBuilderService.TaskPaneBuildResult> renderHost)
        {
            HostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
            FormatContextDescriptor = formatContextDescriptor ?? throw new ArgumentNullException(nameof(formatContextDescriptor));
            FormatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
            FormatWorkbookDescriptor = formatWorkbookDescriptor ?? throw new ArgumentNullException(nameof(formatWorkbookDescriptor));
            RenderHost = renderHost ?? throw new ArgumentNullException(nameof(renderHost));
        }

        internal Dictionary<string, TaskPaneHost> HostsByWindowKey { get; }

        internal Func<WorkbookContext, string> FormatContextDescriptor { get; }

        internal Func<TaskPaneHost, string> FormatHostDescriptor { get; }

        internal Func<Excel.Workbook, string> FormatWorkbookDescriptor { get; }

        internal Func<TaskPaneHost, WorkbookContext, string, TaskPaneSnapshotBuilderService.TaskPaneBuildResult> RenderHost { get; }
    }

    // Create-side adapter compose input. This narrows host-factory wiring so the helper does not receive the
    // whole runtime graph surface/context when it only needs VSTO adapter construction input.
    internal sealed class TaskPaneHostFactoryComposeContext
    {
        internal TaskPaneHostFactoryComposeContext(
            ThisAddIn addIn,
            Logger logger,
            Func<TaskPaneHost, string> formatHostDescriptor)
        {
            AddIn = addIn;
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            FormatHostDescriptor = formatHostDescriptor ?? throw new ArgumentNullException(nameof(formatHostDescriptor));
        }

        internal ThisAddIn AddIn { get; }

        internal Logger Logger { get; }

        internal Func<TaskPaneHost, string> FormatHostDescriptor { get; }
    }

    // Registry-side adapter compose input. This keeps shared-map and diagnostic formatter ownership explicit
    // without routing the full graph context through registry helper compose.
    internal sealed class TaskPaneHostRegistryComposeContext
    {
        internal TaskPaneHostRegistryComposeContext(
            Dictionary<string, TaskPaneHost> hostsByWindowKey,
            Logger logger,
            Func<TaskPaneHost, string> formatHostDescriptorForDiagnostics)
        {
            HostsByWindowKey = hostsByWindowKey ?? throw new ArgumentNullException(nameof(hostsByWindowKey));
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            FormatHostDescriptorForDiagnostics = formatHostDescriptorForDiagnostics ?? throw new ArgumentNullException(nameof(formatHostDescriptorForDiagnostics));
        }

        internal Dictionary<string, TaskPaneHost> HostsByWindowKey { get; }

        internal Logger Logger { get; }

        internal Func<TaskPaneHost, string> FormatHostDescriptorForDiagnostics { get; }
    }

    // Non-case action helper compose input. This keeps Kernel/Accounting action wiring explicit without
    // routing the full runtime graph context through the non-case action helper compose.
    internal sealed class TaskPaneNonCaseActionHandlerComposeContext
    {
        internal TaskPaneNonCaseActionHandlerComposeContext(
            ExcelInteropService excelInteropService,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            AccountingInternalCommandService accountingInternalCommandService,
            UserErrorService userErrorService,
            Logger logger,
            Func<string, TaskPaneHost> resolveHost,
            Action<TaskPaneHost, WorkbookContext, string> renderHost,
            Func<TaskPaneHost, string, bool> tryShowHost)
        {
            ExcelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            KernelCommandService = kernelCommandService ?? throw new ArgumentNullException(nameof(kernelCommandService));
            AccountingSheetCommandService = accountingSheetCommandService ?? throw new ArgumentNullException(nameof(accountingSheetCommandService));
            AccountingInternalCommandService = accountingInternalCommandService ?? throw new ArgumentNullException(nameof(accountingInternalCommandService));
            UserErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            ResolveHost = resolveHost ?? throw new ArgumentNullException(nameof(resolveHost));
            RenderHost = renderHost ?? throw new ArgumentNullException(nameof(renderHost));
            TryShowHost = tryShowHost ?? throw new ArgumentNullException(nameof(tryShowHost));
        }

        internal ExcelInteropService ExcelInteropService { get; }

        internal KernelCommandService KernelCommandService { get; }

        internal AccountingSheetCommandService AccountingSheetCommandService { get; }

        internal AccountingInternalCommandService AccountingInternalCommandService { get; }

        internal UserErrorService UserErrorService { get; }

        internal Logger Logger { get; }

        internal Func<string, TaskPaneHost> ResolveHost { get; }

        internal Action<TaskPaneHost, WorkbookContext, string> RenderHost { get; }

        internal Func<TaskPaneHost, string, bool> TryShowHost { get; }
    }

    // CASE target-resolver compose input. This keeps host lookup wiring explicit without routing the full
    // runtime graph context through the target-resolution helper compose.
    internal sealed class TaskPaneCaseActionTargetResolverComposeContext
    {
        internal TaskPaneCaseActionTargetResolverComposeContext(
            ExcelInteropService excelInteropService,
            Logger logger,
            Func<string, TaskPaneHost> resolveHost)
        {
            ExcelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            ResolveHost = resolveHost ?? throw new ArgumentNullException(nameof(resolveHost));
        }

        internal ExcelInteropService ExcelInteropService { get; }

        internal Logger Logger { get; }

        internal Func<string, TaskPaneHost> ResolveHost { get; }
    }

    // CASE separated-action handler compose input. Both document/accounting handlers share the same collaborator
    // set, so the compose helper keeps that payload explicit without reusing the full graph context.
    internal sealed class TaskPaneCaseActionHandlerComposeContext
    {
        internal TaskPaneCaseActionHandlerComposeContext(
            TaskPaneCaseActionTargetResolver caseActionTargetResolver,
            TaskPaneCaseFallbackActionExecutor taskPaneCaseFallbackActionExecutor,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh)
        {
            CaseActionTargetResolver = caseActionTargetResolver ?? throw new ArgumentNullException(nameof(caseActionTargetResolver));
            TaskPaneCaseFallbackActionExecutor = taskPaneCaseFallbackActionExecutor ?? throw new ArgumentNullException(nameof(taskPaneCaseFallbackActionExecutor));
            CaseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            UserErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            HandlePostActionRefresh = handlePostActionRefresh ?? throw new ArgumentNullException(nameof(handlePostActionRefresh));
        }

        internal TaskPaneCaseActionTargetResolver CaseActionTargetResolver { get; }

        internal TaskPaneCaseFallbackActionExecutor TaskPaneCaseFallbackActionExecutor { get; }

        internal CaseTaskPaneViewStateBuilder CaseTaskPaneViewStateBuilder { get; }

        internal UserErrorService UserErrorService { get; }

        internal Logger Logger { get; }

        internal Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> HandlePostActionRefresh { get; }
    }

    // CASE dispatcher compose input. This keeps the dispatcher-side wiring explicit after the separated handler
    // subtree is resolved, without routing the full runtime graph context through dispatcher compose.
    internal sealed class TaskPaneActionDispatcherComposeContext
    {
        internal TaskPaneActionDispatcherComposeContext(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneCaseFallbackActionExecutor taskPaneCaseFallbackActionExecutor,
            TaskPaneCaseActionTargetResolver caseActionTargetResolver,
            TaskPaneCaseAccountingActionHandler taskPaneCaseAccountingActionHandler,
            TaskPaneCaseDocumentActionHandler taskPaneCaseDocumentActionHandler,
            Action<TaskPaneHost> invalidateHostRenderStateForForcedRefresh,
            Action<DocumentButtonsControl, Excel.Workbook> renderCaseHostAfterAction,
            Func<TaskPaneHost, string, bool> tryShowHost)
        {
            AddIn = addIn;
            ExcelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            CaseTaskPaneViewStateBuilder = caseTaskPaneViewStateBuilder ?? throw new ArgumentNullException(nameof(caseTaskPaneViewStateBuilder));
            UserErrorService = userErrorService ?? throw new ArgumentNullException(nameof(userErrorService));
            Logger = logger ?? throw new ArgumentNullException(nameof(logger));
            TaskPaneCaseFallbackActionExecutor = taskPaneCaseFallbackActionExecutor ?? throw new ArgumentNullException(nameof(taskPaneCaseFallbackActionExecutor));
            CaseActionTargetResolver = caseActionTargetResolver ?? throw new ArgumentNullException(nameof(caseActionTargetResolver));
            TaskPaneCaseAccountingActionHandler = taskPaneCaseAccountingActionHandler ?? throw new ArgumentNullException(nameof(taskPaneCaseAccountingActionHandler));
            TaskPaneCaseDocumentActionHandler = taskPaneCaseDocumentActionHandler ?? throw new ArgumentNullException(nameof(taskPaneCaseDocumentActionHandler));
            InvalidateHostRenderStateForForcedRefresh = invalidateHostRenderStateForForcedRefresh ?? throw new ArgumentNullException(nameof(invalidateHostRenderStateForForcedRefresh));
            RenderCaseHostAfterAction = renderCaseHostAfterAction ?? throw new ArgumentNullException(nameof(renderCaseHostAfterAction));
            TryShowHost = tryShowHost ?? throw new ArgumentNullException(nameof(tryShowHost));
        }

        internal ThisAddIn AddIn { get; }

        internal ExcelInteropService ExcelInteropService { get; }

        internal CaseTaskPaneViewStateBuilder CaseTaskPaneViewStateBuilder { get; }

        internal UserErrorService UserErrorService { get; }

        internal Logger Logger { get; }

        internal TaskPaneCaseFallbackActionExecutor TaskPaneCaseFallbackActionExecutor { get; }

        internal TaskPaneCaseActionTargetResolver CaseActionTargetResolver { get; }

        internal TaskPaneCaseAccountingActionHandler TaskPaneCaseAccountingActionHandler { get; }

        internal TaskPaneCaseDocumentActionHandler TaskPaneCaseDocumentActionHandler { get; }

        internal Action<TaskPaneHost> InvalidateHostRenderStateForForcedRefresh { get; }

        internal Action<DocumentButtonsControl, Excel.Workbook> RenderCaseHostAfterAction { get; }

        internal Func<TaskPaneHost, string, bool> TryShowHost { get; }
    }

    internal static class TaskPaneManagerRuntimeBootstrap
    {
        // Production runtime entrypoint. AddInCompositionRoot should use this path so graph build/attach timing
        // stays fixed at one boundary.
        internal static TaskPaneManager CreateAttached(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger)
        {
            return CreateAttachedCore(
                TaskPaneManagerRuntimeEntryContext.CreateProductionFull(
                    addIn,
                    excelInteropService,
                    taskPaneBusinessActionLauncher,
                    kernelCommandService,
                    accountingSheetCommandService,
                    caseTaskPaneViewStateBuilder,
                    casePaneSnapshotRenderService,
                    accountingInternalCommandService,
                    kernelCaseInteractionState,
                    userErrorService,
                    logger));
        }

        // Full attached test/snapshot entrypoint. Keeps production attach timing while allowing hooks
        // and custom snapshot services for harness scenarios.
        internal static TaskPaneManager CreateAttachedForTests(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            TaskPaneBusinessActionLauncher taskPaneBusinessActionLauncher,
            KernelCommandService kernelCommandService,
            AccountingSheetCommandService accountingSheetCommandService,
            CaseTaskPaneViewStateBuilder caseTaskPaneViewStateBuilder,
            CasePaneSnapshotRenderService casePaneSnapshotRenderService,
            AccountingInternalCommandService accountingInternalCommandService,
            KernelCaseInteractionState kernelCaseInteractionState,
            UserErrorService userErrorService,
            Logger logger,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks = null)
        {
            return CreateAttachedCore(
                TaskPaneManagerRuntimeEntryContext.CreateFullForTests(
                    addIn,
                    excelInteropService,
                    taskPaneBusinessActionLauncher,
                    kernelCommandService,
                    accountingSheetCommandService,
                    caseTaskPaneViewStateBuilder,
                    casePaneSnapshotRenderService,
                    accountingInternalCommandService,
                    kernelCaseInteractionState,
                    userErrorService,
                    logger,
                    testHooks));
        }

        // Thin attached test entrypoint for orchestration-only harnesses that intentionally omit the full action graph.
        internal static TaskPaneManager CreateThinAttachedForTests(
            Logger logger,
            KernelCaseInteractionState kernelCaseInteractionState,
            TaskPaneManager.TaskPaneManagerTestHooks testHooks = null)
        {
            return CreateAttachedCore(TaskPaneManagerRuntimeEntryContext.CreateThinForTests(logger, kernelCaseInteractionState, testHooks));
        }

        // Bootstrap owns the attach order: raw manager construction first, graph composition second, attach last.
        private static TaskPaneManager CreateAttachedCore(TaskPaneManagerRuntimeEntryContext entryContext)
        {
            if (entryContext == null)
            {
                throw new ArgumentNullException(nameof(entryContext));
            }

            TaskPaneManager manager = entryContext.UsesFullRuntimeConstruction
                ? TaskPaneManager.RuntimeBootstrapAccess.CreateUnattachedFullForBootstrap(entryContext)
                : TaskPaneManager.RuntimeBootstrapAccess.CreateUnattachedThinForBootstrap(entryContext);
            TaskPaneManagerRuntimeGraphComposeContext graphContext = entryContext.CreateGraphComposeContext();
            TaskPaneManagerRuntimeGraphComposeSurface graphSurface = TaskPaneManager.RuntimeBootstrapAccess.CreateGraphComposeSurfaceForBootstrap(manager);

            TaskPaneManager.RuntimeBootstrapAccess.AttachRuntimeGraphForBootstrap(
                manager,
                TaskPaneManagerRuntimeGraphFactory.Compose(graphSurface, graphContext));
            return manager;
        }
    }

    internal static class TaskPaneManagerRuntimeGraphFactory
    {
        internal static TaskPaneManagerRuntimeGraph Compose(
            TaskPaneManagerRuntimeGraphComposeSurface graphSurface,
            TaskPaneManagerRuntimeGraphComposeContext graphContext)
        {
            if (graphSurface == null)
            {
                throw new ArgumentNullException(nameof(graphSurface));
            }

            if (graphContext == null)
            {
                throw new ArgumentNullException(nameof(graphContext));
            }

            var casePaneCacheRefreshNotificationService = new CasePaneCacheRefreshNotificationService(
                graphContext.Logger,
                workbook => ResolveWorkbookFullName(graphContext, workbook),
                graphContext.TestHooks != null && graphContext.TestHooks.OnCasePaneUpdatedNotification != null
                    ? new Action<string>(reason => graphContext.TestHooks.OnCasePaneUpdatedNotification(reason))
                    : null);

            TaskPaneHostLifecycleService taskPaneHostLifecycleService = null;
            TaskPaneNonCaseActionHandler taskPaneNonCaseActionHandler = null;
            TaskPaneActionDispatcher taskPaneActionDispatcher = null;
            var taskPaneHostFactoryComposeContext = new TaskPaneHostFactoryComposeContext(
                graphContext.AddIn,
                graphContext.Logger,
                graphSurface.FormatHostDescriptor);

            // Create-side adapter composition only:
            // this graph factory decides who collaborates with the create path, but it does not own runtime create timing.
            // Once refresh/lifecycle code calls into GetOrReplaceHost/CreateHost, the concrete timing still belongs to the
            // existing TaskPaneHostRegistry -> TaskPaneHostFactory -> TaskPaneHost -> ThisAddIn chain.
            TaskPaneHostFactory taskPaneHostFactory = CreateTaskPaneHostFactory(
                taskPaneHostFactoryComposeContext,
                () => taskPaneNonCaseActionHandler,
                () => taskPaneActionDispatcher);
            var taskPaneHostRegistryComposeContext = new TaskPaneHostRegistryComposeContext(
                graphSurface.HostsByWindowKey,
                graphContext.Logger,
                graphSurface.FormatHostDescriptor);
            // Compose-time owner only:
            // this graph decides which collaborators are wired into remove-side ownership, but it does not own
            // standard remove, replacement remove, or shutdown cleanup timing after composition completes.
            TaskPaneHostRegistry taskPaneHostRegistry = CreateTaskPaneHostRegistry(taskPaneHostRegistryComposeContext, taskPaneHostFactory);

            var taskPaneDisplayCoordinator = new TaskPaneDisplayCoordinator(
                graphSurface.HostsByWindowKey,
                graphContext.KernelCaseInteractionState,
                graphContext.Logger,
                graphContext.TestHooks,
                TaskPaneManager.SafeGetWindowKey,
                graphSurface.FormatHostDescriptor,
                graphSurface.FormatWorkbookDescriptor,
                TaskPaneManager.FormatWindowDescriptor,
                windowKey => taskPaneHostLifecycleService.RemoveHost(windowKey));

            taskPaneHostLifecycleService = new TaskPaneHostLifecycleService(
                graphSurface.HostsByWindowKey,
                taskPaneHostRegistry,
                graphContext.ExcelInteropService,
                graphContext.Logger);

            Func<string, TaskPaneHost> resolveHost =
                windowKey => graphSurface.HostsByWindowKey.TryGetValue(windowKey ?? string.Empty, out TaskPaneHost host) ? host : null;
            Func<TaskPaneHost, string, bool> tryShowHost =
                (host, reason) => taskPaneDisplayCoordinator.TryShowHost(host, reason);
            Action<TaskPaneHost, WorkbookContext, string> renderHostWithoutFacts =
                (host, context, reason) =>
                {
                    graphSurface.RenderHost(host, context, reason);
                };

            TaskPaneNonCaseActionHandlerComposeContext taskPaneNonCaseActionHandlerComposeContext = null;
            if (graphContext.ExcelInteropService != null
                && graphContext.KernelCommandService != null
                && graphContext.AccountingSheetCommandService != null
                && graphContext.AccountingInternalCommandService != null
                && graphContext.UserErrorService != null)
            {
                taskPaneNonCaseActionHandlerComposeContext = new TaskPaneNonCaseActionHandlerComposeContext(
                    graphContext.ExcelInteropService,
                    graphContext.KernelCommandService,
                    graphContext.AccountingSheetCommandService,
                    graphContext.AccountingInternalCommandService,
                    graphContext.UserErrorService,
                    graphContext.Logger,
                    resolveHost,
                    renderHostWithoutFacts,
                    tryShowHost);
            }

            if (taskPaneNonCaseActionHandlerComposeContext != null)
            {
                taskPaneNonCaseActionHandler = CreateTaskPaneNonCaseActionHandler(taskPaneNonCaseActionHandlerComposeContext);
            }

            if (CanComposeCaseActionDispatcher(graphContext))
            {
                var taskPaneCaseFallbackActionExecutor = new TaskPaneCaseFallbackActionExecutor(graphContext.TaskPaneBusinessActionLauncher);
                var taskPaneCaseActionTargetResolverComposeContext = new TaskPaneCaseActionTargetResolverComposeContext(
                    graphContext.ExcelInteropService,
                    graphContext.Logger,
                    resolveHost);
                TaskPaneCaseActionTargetResolver taskPaneCaseActionTargetResolver =
                    CreateTaskPaneCaseActionTargetResolver(taskPaneCaseActionTargetResolverComposeContext);
                Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh =
                    (host, workbook, control, actionKind) => taskPaneActionDispatcher.HandlePostActionRefresh(host, workbook, control, actionKind);
                var taskPaneCaseActionHandlerComposeContext = new TaskPaneCaseActionHandlerComposeContext(
                    taskPaneCaseActionTargetResolver,
                    taskPaneCaseFallbackActionExecutor,
                    graphContext.CaseTaskPaneViewStateBuilder,
                    graphContext.UserErrorService,
                    graphContext.Logger,
                    handlePostActionRefresh);
                TaskPaneCaseAccountingActionHandler taskPaneCaseAccountingActionHandler =
                    CreateTaskPaneCaseAccountingActionHandler(taskPaneCaseActionHandlerComposeContext);
                TaskPaneCaseDocumentActionHandler taskPaneCaseDocumentActionHandler =
                    CreateTaskPaneCaseDocumentActionHandler(taskPaneCaseActionHandlerComposeContext);
                var taskPaneActionDispatcherComposeContext = new TaskPaneActionDispatcherComposeContext(
                    graphContext.AddIn,
                    graphContext.ExcelInteropService,
                    graphContext.CaseTaskPaneViewStateBuilder,
                    graphContext.UserErrorService,
                    graphContext.Logger,
                    taskPaneCaseFallbackActionExecutor,
                    taskPaneCaseActionTargetResolver,
                    taskPaneCaseAccountingActionHandler,
                    taskPaneCaseDocumentActionHandler,
                    host => taskPaneDisplayCoordinator.InvalidateHostRenderStateForForcedRefresh(host),
                    (control, workbook) => graphContext.CasePaneSnapshotRenderService.RenderAfterAction(control, workbook),
                    tryShowHost);
                taskPaneActionDispatcher = CreateTaskPaneActionDispatcher(taskPaneActionDispatcherComposeContext);
            }

            var taskPaneHostFlowService = new TaskPaneHostFlowService(
                graphContext.ExcelInteropService,
                taskPaneDisplayCoordinator,
                taskPaneHostLifecycleService,
                graphContext.Logger,
                graphSurface.FormatContextDescriptor,
                graphSurface.FormatHostDescriptor,
                TaskPaneManager.SafeGetWindowKey,
                graphSurface.RenderHost);

            return new TaskPaneManagerRuntimeGraph(
                casePaneCacheRefreshNotificationService,
                taskPaneHostLifecycleService,
                taskPaneDisplayCoordinator,
                taskPaneHostFlowService);
        }

        private static TaskPaneHostFactory CreateTaskPaneHostFactory(
            TaskPaneHostFactoryComposeContext composeContext,
            Func<TaskPaneNonCaseActionHandler> resolveNonCaseActionHandler,
            Func<TaskPaneActionDispatcher> resolveCaseActionDispatcher)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            // Compose-time intent only. Runtime ActionInvoked binding timing remains inside TaskPaneHostFactory/CreateHost(...).
            return new TaskPaneHostFactory(
                composeContext.AddIn,
                composeContext.Logger,
                composeContext.FormatHostDescriptor,
                (windowKey, e) => resolveNonCaseActionHandler()?.HandleKernelActionInvoked(windowKey, e),
                (windowKey, e) => resolveNonCaseActionHandler()?.HandleAccountingActionInvoked(windowKey, e),
                (windowKey, control, e) => resolveCaseActionDispatcher()?.HandleCaseControlActionInvoked(windowKey, control, e));
        }

        private static TaskPaneHostRegistry CreateTaskPaneHostRegistry(
            TaskPaneHostRegistryComposeContext composeContext,
            TaskPaneHostFactory taskPaneHostFactory)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            return new TaskPaneHostRegistry(
                composeContext.HostsByWindowKey,
                composeContext.Logger,
                // Diagnostic-only input for remove-host trace output. Registry does not own identity or metadata timing through this formatter.
                composeContext.FormatHostDescriptorForDiagnostics,
                taskPaneHostFactory);
        }

        private static TaskPaneNonCaseActionHandler CreateTaskPaneNonCaseActionHandler(
            TaskPaneNonCaseActionHandlerComposeContext composeContext)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            return new TaskPaneNonCaseActionHandler(
                composeContext.ExcelInteropService,
                composeContext.KernelCommandService,
                composeContext.AccountingSheetCommandService,
                composeContext.AccountingInternalCommandService,
                composeContext.UserErrorService,
                composeContext.Logger,
                composeContext.ResolveHost,
                composeContext.RenderHost,
                composeContext.TryShowHost);
        }

        private static TaskPaneCaseActionTargetResolver CreateTaskPaneCaseActionTargetResolver(
            TaskPaneCaseActionTargetResolverComposeContext composeContext)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            return new TaskPaneCaseActionTargetResolver(
                composeContext.ExcelInteropService,
                composeContext.Logger,
                composeContext.ResolveHost);
        }

        private static TaskPaneCaseAccountingActionHandler CreateTaskPaneCaseAccountingActionHandler(
            TaskPaneCaseActionHandlerComposeContext composeContext)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            return new TaskPaneCaseAccountingActionHandler(
                composeContext.CaseActionTargetResolver,
                composeContext.TaskPaneCaseFallbackActionExecutor,
                composeContext.CaseTaskPaneViewStateBuilder,
                composeContext.UserErrorService,
                composeContext.Logger,
                composeContext.HandlePostActionRefresh);
        }

        private static TaskPaneCaseDocumentActionHandler CreateTaskPaneCaseDocumentActionHandler(
            TaskPaneCaseActionHandlerComposeContext composeContext)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            return new TaskPaneCaseDocumentActionHandler(
                composeContext.CaseActionTargetResolver,
                composeContext.TaskPaneCaseFallbackActionExecutor,
                composeContext.CaseTaskPaneViewStateBuilder,
                composeContext.UserErrorService,
                composeContext.Logger,
                composeContext.HandlePostActionRefresh);
        }

        private static TaskPaneActionDispatcher CreateTaskPaneActionDispatcher(
            TaskPaneActionDispatcherComposeContext composeContext)
        {
            if (composeContext == null)
            {
                throw new ArgumentNullException(nameof(composeContext));
            }

            return new TaskPaneActionDispatcher(
                composeContext.AddIn,
                composeContext.ExcelInteropService,
                composeContext.CaseTaskPaneViewStateBuilder,
                composeContext.UserErrorService,
                composeContext.Logger,
                composeContext.TaskPaneCaseFallbackActionExecutor,
                composeContext.CaseActionTargetResolver,
                composeContext.TaskPaneCaseAccountingActionHandler,
                composeContext.TaskPaneCaseDocumentActionHandler,
                composeContext.InvalidateHostRenderStateForForcedRefresh,
                composeContext.RenderCaseHostAfterAction,
                composeContext.TryShowHost);
        }

        private static string ResolveWorkbookFullName(TaskPaneManagerRuntimeGraphComposeContext graphContext, Excel.Workbook workbook)
        {
            if (graphContext.ExcelInteropService != null)
            {
                return graphContext.ExcelInteropService.GetWorkbookFullName(workbook) ?? string.Empty;
            }

            return workbook == null ? string.Empty : (workbook.FullName ?? string.Empty);
        }

        private static bool CanComposeCaseActionDispatcher(TaskPaneManagerRuntimeGraphComposeContext graphContext)
        {
            return graphContext.ExcelInteropService != null
                && graphContext.TaskPaneBusinessActionLauncher != null
                && graphContext.CaseTaskPaneViewStateBuilder != null
                && graphContext.CasePaneSnapshotRenderService != null
                && graphContext.UserErrorService != null;
        }
    }
}
