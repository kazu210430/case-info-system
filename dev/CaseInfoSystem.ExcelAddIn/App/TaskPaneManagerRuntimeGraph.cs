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
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
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
            CaseTaskPaneSnapshotReader = caseTaskPaneSnapshotReader;
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
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
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
                caseTaskPaneSnapshotReader,
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
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
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
                caseTaskPaneSnapshotReader ?? throw new ArgumentNullException(nameof(caseTaskPaneSnapshotReader)),
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
                caseTaskPaneSnapshotReader: null,
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

        internal ICaseTaskPaneSnapshotReader CaseTaskPaneSnapshotReader { get; }

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
            Action<TaskPaneHost, WorkbookContext, string> renderHost)
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

        internal Action<TaskPaneHost, WorkbookContext, string> RenderHost { get; }
    }

    internal static class TaskPaneManagerRuntimeBootstrap
    {
        // Production runtime entrypoint. AddInCompositionRoot should use this path so graph build/attach timing
        // stays fixed at one boundary.
        internal static TaskPaneManager CreateAttached(
            ThisAddIn addIn,
            ExcelInteropService excelInteropService,
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
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
                    caseTaskPaneSnapshotReader,
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
            ICaseTaskPaneSnapshotReader caseTaskPaneSnapshotReader,
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
                    caseTaskPaneSnapshotReader,
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

            // Create-side adapter composition only:
            // this graph factory decides who collaborates with the create path, but it does not own runtime create timing.
            // Once refresh/lifecycle code calls into GetOrReplaceHost/CreateHost, the concrete timing still belongs to the
            // existing TaskPaneHostRegistry -> TaskPaneHostFactory -> TaskPaneHost -> ThisAddIn chain.
            TaskPaneHostFactory taskPaneHostFactory = CreateTaskPaneHostFactory(
                graphSurface,
                graphContext,
                () => taskPaneNonCaseActionHandler,
                () => taskPaneActionDispatcher);
            // Compose-time owner only:
            // this graph decides which collaborators are wired into remove-side ownership, but it does not own
            // standard remove, replacement remove, or shutdown cleanup timing after composition completes.
            TaskPaneHostRegistry taskPaneHostRegistry = CreateTaskPaneHostRegistry(graphSurface, graphContext, taskPaneHostFactory);

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

            if (CanComposeNonCaseActionHandler(graphContext))
            {
                taskPaneNonCaseActionHandler = new TaskPaneNonCaseActionHandler(
                    graphContext.ExcelInteropService,
                    graphContext.KernelCommandService,
                    graphContext.AccountingSheetCommandService,
                    graphContext.AccountingInternalCommandService,
                    graphContext.UserErrorService,
                    graphContext.Logger,
                    resolveHost,
                    graphSurface.RenderHost,
                    (host, reason) => taskPaneDisplayCoordinator.TryShowHost(host, reason));
            }

            if (CanComposeCaseActionDispatcher(graphContext))
            {
                var taskPaneCaseFallbackActionExecutor = new TaskPaneCaseFallbackActionExecutor(graphContext.TaskPaneBusinessActionLauncher);
                var taskPaneCaseActionTargetResolver = new TaskPaneCaseActionTargetResolver(
                    graphContext.ExcelInteropService,
                    graphContext.Logger,
                    resolveHost);
                Action<TaskPaneHost, Excel.Workbook, DocumentButtonsControl, string> handlePostActionRefresh =
                    (host, workbook, control, actionKind) => taskPaneActionDispatcher.HandlePostActionRefresh(host, workbook, control, actionKind);
                var taskPaneCaseAccountingActionHandler = new TaskPaneCaseAccountingActionHandler(
                    taskPaneCaseActionTargetResolver,
                    taskPaneCaseFallbackActionExecutor,
                    graphContext.CaseTaskPaneViewStateBuilder,
                    graphContext.UserErrorService,
                    graphContext.Logger,
                    handlePostActionRefresh);
                var taskPaneCaseDocumentActionHandler = new TaskPaneCaseDocumentActionHandler(
                    taskPaneCaseActionTargetResolver,
                    taskPaneCaseFallbackActionExecutor,
                    graphContext.CaseTaskPaneViewStateBuilder,
                    graphContext.UserErrorService,
                    graphContext.Logger,
                    handlePostActionRefresh);

                taskPaneActionDispatcher = new TaskPaneActionDispatcher(
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
                    (host, reason) => taskPaneDisplayCoordinator.TryShowHost(host, reason));
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
            TaskPaneManagerRuntimeGraphComposeSurface graphSurface,
            TaskPaneManagerRuntimeGraphComposeContext graphContext,
            Func<TaskPaneNonCaseActionHandler> resolveNonCaseActionHandler,
            Func<TaskPaneActionDispatcher> resolveCaseActionDispatcher)
        {
            // Compose-time intent only. Runtime ActionInvoked binding timing remains inside TaskPaneHostFactory/CreateHost(...).
            return new TaskPaneHostFactory(
                graphContext.AddIn,
                graphContext.Logger,
                graphSurface.FormatHostDescriptor,
                (windowKey, e) => resolveNonCaseActionHandler()?.HandleKernelActionInvoked(windowKey, e),
                (windowKey, e) => resolveNonCaseActionHandler()?.HandleAccountingActionInvoked(windowKey, e),
                (windowKey, control, e) => resolveCaseActionDispatcher()?.HandleCaseControlActionInvoked(windowKey, control, e));
        }

        private static TaskPaneHostRegistry CreateTaskPaneHostRegistry(
            TaskPaneManagerRuntimeGraphComposeSurface graphSurface,
            TaskPaneManagerRuntimeGraphComposeContext graphContext,
            TaskPaneHostFactory taskPaneHostFactory)
        {
            return new TaskPaneHostRegistry(
                graphSurface.HostsByWindowKey,
                graphContext.Logger,
                // Diagnostic-only input for remove-host trace output. Registry does not own identity or metadata timing through this formatter.
                graphSurface.FormatHostDescriptor,
                taskPaneHostFactory);
        }

        private static string ResolveWorkbookFullName(TaskPaneManagerRuntimeGraphComposeContext graphContext, Excel.Workbook workbook)
        {
            if (graphContext.ExcelInteropService != null)
            {
                return graphContext.ExcelInteropService.GetWorkbookFullName(workbook) ?? string.Empty;
            }

            return workbook == null ? string.Empty : (workbook.FullName ?? string.Empty);
        }

        private static bool CanComposeNonCaseActionHandler(TaskPaneManagerRuntimeGraphComposeContext graphContext)
        {
            return graphContext.ExcelInteropService != null
                && graphContext.KernelCommandService != null
                && graphContext.AccountingSheetCommandService != null
                && graphContext.AccountingInternalCommandService != null
                && graphContext.UserErrorService != null;
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
