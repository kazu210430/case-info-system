using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class KernelUserDataRegistrationExecutionServiceTests
    {
        [Fact]
        public void Execute_ShowsWaitUiAndClosesItAfterSuccessfulReflection()
        {
            var reflectionService = new FakeKernelUserDataReflectionService();
            var waitService = new FakeKernelUserDataRegistrationWaitService();
            var service = CreateService(reflectionService, waitService);
            var context = new WorkbookContext(null, null, WorkbookRole.Kernel, @"C:\root", @"C:\root\Kernel.xlsx", "shUserData");

            service.Execute(context);

            Assert.Same(context, reflectionService.LastReflectAllContext);
            Assert.Equal(1, waitService.ShowWaitingCallCount);
            Assert.True(waitService.LastSessionDisposed);
        }

        [Fact]
        public void Execute_WhenReflectionThrows_ClosesWaitUiBeforeRethrow()
        {
            var reflectionService = new FakeKernelUserDataReflectionService
            {
                ReflectAllException = new InvalidOperationException("boom")
            };
            var waitService = new FakeKernelUserDataRegistrationWaitService();
            var service = CreateService(reflectionService, waitService);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                service.Execute(new WorkbookContext(null, null, WorkbookRole.Kernel, @"C:\root", @"C:\root\Kernel.xlsx", "shUserData")));

            Assert.Equal("boom", exception.Message);
            Assert.Equal(1, waitService.ShowWaitingCallCount);
            Assert.True(waitService.LastSessionDisposed);
        }

        private static KernelUserDataRegistrationExecutionService CreateService(
            IKernelUserDataReflectionService reflectionService,
            IKernelUserDataRegistrationWaitService waitService)
        {
            return new KernelUserDataRegistrationExecutionService(
                reflectionService,
                waitService,
                new Logger(_ => { }));
        }

        private sealed class FakeKernelUserDataReflectionService : IKernelUserDataReflectionService
        {
            internal WorkbookContext LastReflectAllContext { get; private set; }

            internal Exception ReflectAllException { get; set; }

            public void ReflectAll(WorkbookContext context)
            {
                LastReflectAllContext = context;
                if (ReflectAllException != null)
                {
                    throw ReflectAllException;
                }
            }

            public void ReflectToAccountingSetOnly(WorkbookContext context)
            {
            }

            public void ReflectToBaseHomeOnly(WorkbookContext context)
            {
            }
        }

        private sealed class FakeKernelUserDataRegistrationWaitService : IKernelUserDataRegistrationWaitService
        {
            internal int ShowWaitingCallCount { get; private set; }

            internal bool LastSessionDisposed { get; private set; }

            public IDisposable ShowWaiting(Stopwatch commandStopwatch)
            {
                ShowWaitingCallCount++;
                LastSessionDisposed = false;
                return new FakeWaitSession(this);
            }

            private sealed class FakeWaitSession : IDisposable
            {
                private readonly FakeKernelUserDataRegistrationWaitService _owner;
                private bool _disposed;

                internal FakeWaitSession(FakeKernelUserDataRegistrationWaitService owner)
                {
                    _owner = owner;
                }

                public void Dispose()
                {
                    if (_disposed)
                    {
                        return;
                    }

                    _disposed = true;
                    _owner.LastSessionDisposed = true;
                }
            }
        }
    }
}
