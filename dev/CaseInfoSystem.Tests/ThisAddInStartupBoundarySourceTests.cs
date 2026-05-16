using System;
using System.IO;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ThisAddInStartupBoundarySourceTests
    {
        [Fact]
        public void AddInStartupBoundaryCoordinator_DoesNotOwnApplicationEventSubscription()
        {
            string source = ReadAddInSource("App", "AddInStartupBoundaryCoordinator.cs");

            Assert.DoesNotContain("ApplicationEventSubscriptionService", source);
            Assert.DoesNotContain("HookApplicationEvents", source);
            Assert.DoesNotContain("UnhookApplicationEvents", source);
            Assert.DoesNotContain("Application_WorkbookOpen", source);
            Assert.DoesNotContain("Application_WorkbookActivate", source);
            Assert.DoesNotContain("Application_WindowActivate", source);
            Assert.DoesNotContain("WorkbookOpen +=", source);
            Assert.DoesNotContain("WorkbookActivate +=", source);
            Assert.DoesNotContain("WindowActivate +=", source);
        }

        [Fact]
        public void ThisAddIn_RetainsVstoEventAndApplicationWiringOwner()
        {
            string source = ReadAddInSource("ThisAddIn.cs");

            Assert.Contains("private void ThisAddIn_Startup", source);
            Assert.Contains("private void ThisAddIn_Shutdown", source);
            Assert.Contains("private void InitializeApplicationEventSubscriptionService", source);
            Assert.Contains("private void HookApplicationEvents", source);
            Assert.Contains("private void UnhookApplicationEvents", source);
            Assert.Contains("private void Application_WorkbookOpen", source);
            Assert.Contains("private void Application_WindowActivate", source);
        }

        [Fact]
        public void ThisAddIn_NoLongerContainsStartupGuardAndExecutionBridgeBodies()
        {
            string thisAddInSource = ReadAddInSource("ThisAddIn.cs");
            string startupCoordinatorSource = ReadAddInSource("App", "AddInStartupBoundaryCoordinator.cs");
            string executionCoordinatorSource = ReadAddInSource("App", "AddInExecutionBoundaryCoordinator.cs");

            Assert.DoesNotContain("private void TraceAndScheduleManagedCloseStartupGuard", thisAddInSource);
            Assert.DoesNotContain("private void QuitEmptyStartupExcelForManagedClose", thisAddInSource);
            Assert.DoesNotContain("private sealed class ManagedCloseStartupFacts", thisAddInSource);
            Assert.DoesNotContain("private void RunWithScreenUpdatingSuspended", thisAddInSource);
            Assert.DoesNotContain("private IDisposable SuppressTaskPaneRefresh", thisAddInSource);
            Assert.DoesNotContain("_taskPaneRefreshSuppressionCount", thisAddInSource);

            Assert.Contains("internal void TraceAndScheduleManagedCloseStartupGuard", startupCoordinatorSource);
            Assert.Contains("private void QuitEmptyStartupExcelForManagedClose", startupCoordinatorSource);
            Assert.Contains("internal sealed class ManagedCloseStartupFacts", startupCoordinatorSource);
            Assert.Contains("_application.ScreenUpdating = false", executionCoordinatorSource);
            Assert.Contains("TaskPaneRefreshSuppressionCount", executionCoordinatorSource);
        }

        private static string ReadAddInSource(params string[] pathParts)
        {
            string repoRoot = FindRepositoryRoot();
            string[] fullPathParts = new string[pathParts.Length + 2];
            fullPathParts[0] = repoRoot;
            fullPathParts[1] = Path.Combine("dev", "CaseInfoSystem.ExcelAddIn");
            Array.Copy(pathParts, 0, fullPathParts, 2, pathParts.Length);
            return File.ReadAllText(Path.Combine(fullPathParts));
        }

        private static string FindRepositoryRoot()
        {
            DirectoryInfo current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "build.ps1"))
                    && Directory.Exists(Path.Combine(current.FullName, "dev", "CaseInfoSystem.ExcelAddIn")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Repository root was not found.");
        }
    }
}
