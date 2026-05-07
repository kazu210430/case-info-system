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
                    context.ExcelInteropService,
                    context.CasePaneSnapshotRenderService,
                    context.Logger);
            }

            internal static TaskPaneManager CreateUnattachedThinForBootstrap(TaskPaneManagerRuntimeEntryContext context)
            {
                if (context == null)
                {
                    throw new ArgumentNullException(nameof(context));
                }

                return new TaskPaneManager(
                    context.Logger);
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
