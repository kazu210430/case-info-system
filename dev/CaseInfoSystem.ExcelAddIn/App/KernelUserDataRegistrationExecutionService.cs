using System;
using System.Diagnostics;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal interface IKernelUserDataRegistrationWaitService
    {
        IDisposable ShowWaiting(Stopwatch commandStopwatch);
    }

    internal sealed class KernelUserDataRegistrationExecutionService
    {
        private readonly IKernelUserDataReflectionService _kernelUserDataReflectionService;
        private readonly IKernelUserDataRegistrationWaitService _waitService;
        private readonly Logger _logger;

        internal KernelUserDataRegistrationExecutionService(
            IKernelUserDataReflectionService kernelUserDataReflectionService,
            IKernelUserDataRegistrationWaitService waitService,
            Logger logger)
        {
            _kernelUserDataReflectionService = kernelUserDataReflectionService ?? throw new ArgumentNullException(nameof(kernelUserDataReflectionService));
            _waitService = waitService ?? throw new ArgumentNullException(nameof(waitService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal void Execute(WorkbookContext context)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            using (_waitService.ShowWaiting(stopwatch))
            {
                _kernelUserDataReflectionService.ReflectAll(context);
            }

            _logger.Info("Kernel user data registration wait flow completed. elapsedMs=" + stopwatch.ElapsedMilliseconds.ToString());
        }
    }
}
