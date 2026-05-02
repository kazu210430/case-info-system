using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class ManagedCloseState
    {
        private readonly Dictionary<string, int> _managedCloseCounts =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        internal IDisposable BeginScope(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return NoOpDisposable.Instance;
            }

            if (_managedCloseCounts.ContainsKey(workbookKey))
            {
                _managedCloseCounts[workbookKey] = _managedCloseCounts[workbookKey] + 1;
            }
            else
            {
                _managedCloseCounts.Add(workbookKey, 1);
            }

            return new ManagedCloseScope(this, workbookKey);
        }

        internal bool IsManagedClose(string workbookKey)
        {
            return workbookKey.Length > 0
                && _managedCloseCounts.TryGetValue(workbookKey, out int count)
                && count > 0;
        }

        internal void Remove(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            _managedCloseCounts.Remove(workbookKey);
        }

        private void Release(string workbookKey)
        {
            if (string.IsNullOrWhiteSpace(workbookKey) || !_managedCloseCounts.TryGetValue(workbookKey, out int count))
            {
                return;
            }

            if (count <= 1)
            {
                _managedCloseCounts.Remove(workbookKey);
                return;
            }

            _managedCloseCounts[workbookKey] = count - 1;
        }

        private sealed class ManagedCloseScope : IDisposable
        {
            private readonly ManagedCloseState _owner;
            private readonly string _workbookKey;
            private bool _disposed;

            internal ManagedCloseScope(ManagedCloseState owner, string workbookKey)
            {
                _owner = owner;
                _workbookKey = workbookKey;
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _owner.Release(_workbookKey);
            }
        }

        private sealed class NoOpDisposable : IDisposable
        {
            internal static readonly NoOpDisposable Instance = new NoOpDisposable();

            public void Dispose()
            {
            }
        }
    }
}
