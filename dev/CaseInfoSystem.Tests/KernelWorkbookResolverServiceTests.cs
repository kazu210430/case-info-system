using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class KernelWorkbookResolverServiceTests
    {
        [Fact]
        public void ResolveOrOpenReadOnly_ReturnsAlreadyOpenWorkbookWithoutTemporaryCloseResponsibility()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            {
                var existingKernelWorkbook = fixture.CreateKernelWorkbook();
                fixture.Application.Workbooks.Add(existingKernelWorkbook);

                KernelWorkbookAccessResult result = fixture.Service.ResolveOrOpenReadOnly(fixture.CreateCaseWorkbook());

                Assert.Same(existingKernelWorkbook, result.Workbook);
                Assert.True(result.WorkbookWasAlreadyOpen);
                Assert.False(result.CallerOwnsTemporaryWorkbook);
                Assert.False(result.WorkbookWasOpenedByResolver);
                Assert.True(existingKernelWorkbook.Windows[1].Visible);

                result.CloseIfOwned("unit-existing-kernel");

                Assert.Equal(0, existingKernelWorkbook.CloseCallCount);
                Assert.Contains(existingKernelWorkbook, fixture.Application.Workbooks);
            }
        }

        [Fact]
        public void ResolveOrOpenReadOnly_OpensTemporaryWorkbookAndAssignsCallerCloseResponsibility()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            {
                Excel.Workbook caseWorkbook = fixture.CreateCaseWorkbook();

                KernelWorkbookAccessResult result = fixture.Service.ResolveOrOpenReadOnly(caseWorkbook);

                Assert.NotNull(result.Workbook);
                Assert.False(result.WorkbookWasAlreadyOpen);
                Assert.True(result.CallerOwnsTemporaryWorkbook);
                Assert.True(result.WorkbookWasOpenedByResolver);
                Assert.Same(fixture.Application, result.Workbook.Application);
                Assert.False(result.Workbook.Windows[1].Visible);

                result.CloseIfOwned("unit-temporary-kernel");

                Assert.Equal(1, result.Workbook.CloseCallCount);
                Assert.False(result.Workbook.LastCloseSaveChanges.GetValueOrDefault());
                Assert.DoesNotContain(result.Workbook, fixture.Application.Workbooks);
            }
        }

        [Fact]
        public void CloseIfOwned_WhenCalledTwice_ClosesTemporaryWorkbookOnlyOnce()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            {
                KernelWorkbookAccessResult result = fixture.Service.ResolveOrOpenReadOnly(fixture.CreateCaseWorkbook());

                result.CloseIfOwned("unit-first-close");
                result.CloseIfOwned("unit-second-close");

                Assert.Equal(1, result.Workbook.CloseCallCount);
            }
        }

        [Fact]
        public void CloseIfOwned_WhenRequested_SuppressesEventsDuringCloseAndRestoresThem()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            {
                fixture.Application.EnableEvents = true;
                KernelWorkbookAccessResult result = fixture.Service.ResolveOrOpenReadOnly(fixture.CreateCaseWorkbook());
                bool eventsSuppressedDuringClose = false;
                result.Workbook.CloseBehavior = () => eventsSuppressedDuringClose = !fixture.Application.EnableEvents;

                result.CloseIfOwned("unit-close-with-events-suppressed", suppressEventsDuringClose: true);

                Assert.True(eventsSuppressedDuringClose);
                Assert.True(fixture.Application.EnableEvents);
            }
        }

        [Fact]
        public void ResolveOrOpenReadOnly_UsesSystemRootPathAndDoesNotTransferOwnershipFromAnotherRoot()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            using (var otherRoot = KernelWorkbookResolverFixture.Create())
            {
                Excel.Workbook otherKernelWorkbook = otherRoot.CreateKernelWorkbook();
                fixture.Application.Workbooks.Add(otherKernelWorkbook);

                KernelWorkbookAccessResult result = fixture.Service.ResolveOrOpenReadOnly(fixture.CreateCaseWorkbook());

                Assert.NotSame(otherKernelWorkbook, result.Workbook);
                Assert.Equal(fixture.KernelWorkbookPath, result.ResolvedKernelPath);
                Assert.True(result.CallerOwnsTemporaryWorkbook);
                Assert.True(result.WorkbookWasOpenedByResolver);

                result.CloseIfOwned("unit-root-bound-kernel");
                Assert.Equal(0, otherKernelWorkbook.CloseCallCount);
            }
        }

        [Fact]
        public void ResolveOrOpenReadOnly_WhenSystemRootIsMissing_FailsClosedWithoutOpeningWorkbook()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            {
                var caseWorkbook = new Excel.Workbook
                {
                    CustomDocumentProperties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                };

                KernelWorkbookAccessResult result = fixture.Service.ResolveOrOpenReadOnly(caseWorkbook);

                Assert.Null(result.Workbook);
                Assert.False(result.CallerOwnsTemporaryWorkbook);
                Assert.False(result.WorkbookWasOpenedByResolver);
                result.CloseIfOwned("unit-missing-root");
                Assert.Empty(fixture.Application.Workbooks);
            }
        }

        [Fact]
        public void ResolveOrOpenReadOnly_WhenOpenFails_RestoresEventsAndPropagates()
        {
            using (var fixture = KernelWorkbookResolverFixture.Create())
            {
                fixture.Application.EnableEvents = true;
                fixture.Application.Workbooks.OpenBehavior = (_, __, ___) => throw new InvalidOperationException("open failed");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => fixture.Service.ResolveOrOpenReadOnly(fixture.CreateCaseWorkbook()));

                Assert.Equal("open failed", exception.Message);
                Assert.True(fixture.Application.EnableEvents);
                Assert.Empty(fixture.Application.Workbooks);
            }
        }

        private sealed class KernelWorkbookResolverFixture : IDisposable
        {
            private KernelWorkbookResolverFixture(string rootPath)
            {
                RootPath = rootPath;
                KernelWorkbookPath = Path.Combine(rootPath, WorkbookFileNameResolver.BuildKernelWorkbookName(".xlsx"));
                File.WriteAllText(KernelWorkbookPath, string.Empty);
                Application = new Excel.Application();
                Logger = new Logger(_ => { });
                PathCompatibilityService = new PathCompatibilityService(Logger);
                ExcelInteropService = new ExcelInteropService(Application, Logger, PathCompatibilityService);
                ExcelInteropService.OnFindOpenWorkbook = FindOpenWorkbook;
                Service = new KernelWorkbookResolverService(Application, ExcelInteropService, PathCompatibilityService, Logger);
            }

            internal string RootPath { get; }

            internal string KernelWorkbookPath { get; }

            internal Excel.Application Application { get; }

            internal Logger Logger { get; }

            internal PathCompatibilityService PathCompatibilityService { get; }

            internal ExcelInteropService ExcelInteropService { get; }

            internal KernelWorkbookResolverService Service { get; }

            internal static KernelWorkbookResolverFixture Create()
            {
                string rootPath = Path.Combine(Path.GetTempPath(), "CaseInfoSystemTests", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(rootPath);
                return new KernelWorkbookResolverFixture(rootPath);
            }

            internal Excel.Workbook CreateCaseWorkbook()
            {
                return new Excel.Workbook
                {
                    FullName = Path.Combine(RootPath, "Case.xlsx"),
                    Name = "Case.xlsx",
                    Path = RootPath,
                    CustomDocumentProperties = CreateDocumentProperties(RootPath)
                };
            }

            internal Excel.Workbook CreateKernelWorkbook()
            {
                return new Excel.Workbook
                {
                    FullName = KernelWorkbookPath,
                    Name = Path.GetFileName(KernelWorkbookPath),
                    Path = RootPath,
                    CustomDocumentProperties = CreateDocumentProperties(RootPath)
                };
            }

            public void Dispose()
            {
                try
                {
                    Directory.Delete(RootPath, recursive: true);
                }
                catch
                {
                }
            }

            private Excel.Workbook FindOpenWorkbook(string workbookPath)
            {
                string normalizedTarget = PathCompatibilityService.NormalizePath(workbookPath);
                foreach (Excel.Workbook workbook in Application.Workbooks)
                {
                    if (workbook == null)
                    {
                        continue;
                    }

                    string workbookFullName = PathCompatibilityService.NormalizePath(workbook.FullName);
                    if (string.Equals(workbookFullName, normalizedTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        return workbook;
                    }
                }

                return null;
            }

            private static IDictionary<string, string> CreateDocumentProperties(string systemRoot)
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SYSTEM_ROOT"] = systemRoot
                };
            }
        }
    }
}
