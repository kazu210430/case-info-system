using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed partial class TaskPaneManager
    {
        // Bridge used only by TaskPaneManagerRuntimeBootstrap. It keeps bootstrap-only construction/attach
        // access away from the core runtime file without changing raw manager construction -> graph compose ->
        // attach ordering.
        internal static class RuntimeBootstrapAccess
        {
            internal static TaskPaneManager CreateUnattachedFullForBootstrap(TaskPaneManagerRuntimeEntryContext context)
            {
                if (context == null)
                {
                    throw new ArgumentNullException(nameof(context));
                }

                return new TaskPaneManager(
                    context.AddIn,
                    context.ExcelInteropService,
                    context.CaseTaskPaneSnapshotReader,
                    context.TaskPaneBusinessActionLauncher,
                    context.KernelCommandService,
                    context.AccountingSheetCommandService,
                    context.CaseTaskPaneViewStateBuilder,
                    context.CasePaneSnapshotRenderService,
                    context.AccountingInternalCommandService,
                    context.KernelCaseInteractionState,
                    context.UserErrorService,
                    context.Logger,
                    context.TestHooks);
            }

            internal static TaskPaneManager CreateUnattachedThinForBootstrap(TaskPaneManagerRuntimeEntryContext context)
            {
                if (context == null)
                {
                    throw new ArgumentNullException(nameof(context));
                }

                return new TaskPaneManager(
                    context.Logger,
                    context.KernelCaseInteractionState,
                    context.TestHooks);
            }

            internal static void AttachRuntimeGraphForBootstrap(TaskPaneManager manager, TaskPaneManagerRuntimeGraph runtimeGraph)
            {
                if (manager == null)
                {
                    throw new ArgumentNullException(nameof(manager));
                }

                manager.AttachRuntimeGraph(runtimeGraph);
            }
        }
    }
}
