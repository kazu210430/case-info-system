using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class WorkbookResetDefinition
    {
        /// <summary>
        internal WorkbookResetDefinition(
            string targetName,
            IReadOnlyDictionary<string, string> fixedValues,
            IReadOnlyList<string> clearPrefixNames)
        {
            if (string.IsNullOrWhiteSpace(targetName))
            {
                throw new ArgumentException("Target name is required.", nameof(targetName));
            }

            TargetName = targetName;
            FixedValues = fixedValues ?? throw new ArgumentNullException(nameof(fixedValues));
            ClearPrefixNames = clearPrefixNames ?? throw new ArgumentNullException(nameof(clearPrefixNames));
        }

        internal string TargetName { get; }

        internal IReadOnlyDictionary<string, string> FixedValues { get; }

        internal IReadOnlyList<string> ClearPrefixNames { get; }
    }
}
