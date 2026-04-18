using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class KernelCaseInteractionState
    {
        private readonly Logger _logger;
        private int _kernelCaseCreationFlowCount;

        /// <summary>
        internal KernelCaseInteractionState(Logger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal IDisposable BeginKernelCaseCreationFlow(string reason)
        {
            _kernelCaseCreationFlowCount++;
            _logger.Info(
                "Kernel case interaction flow started. reason="
                + (reason ?? string.Empty)
                + ", count="
                + _kernelCaseCreationFlowCount.ToString());
            return new Scope(this, reason);
        }

        /// <summary>
        internal bool IsKernelCaseCreationFlowActive
        {
            get { return _kernelCaseCreationFlowCount > 0; }
        }

        /// <summary>
        private void EndKernelCaseCreationFlow(string reason)
        {
            if (_kernelCaseCreationFlowCount > 0)
            {
                _kernelCaseCreationFlowCount--;
            }

            _logger.Info(
                "Kernel case interaction flow ended. reason="
                + (reason ?? string.Empty)
                + ", count="
                + _kernelCaseCreationFlowCount.ToString());
        }

        /// <summary>
        private sealed class Scope : IDisposable
        {
            private readonly KernelCaseInteractionState _owner;
            private readonly string _reason;
            private bool _disposed;

            /// <summary>
            internal Scope(KernelCaseInteractionState owner, string reason)
            {
                _owner = owner;
                _reason = reason ?? string.Empty;
            }

            /// <summary>
            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _owner.EndKernelCaseCreationFlow(_reason);
            }
        }
    }
}
