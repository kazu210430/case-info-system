using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class WorkbookResetDefinitionRepository
    {
        private static readonly IReadOnlyDictionary<string, string> KernelFixedValues =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SYSTEM_ROOT", string.Empty },
                { "WORD_TEMPLATE_DIR", string.Empty },
                { "DEFAULT_ROOT", string.Empty },
                { "LAST_PICK_FOLDER", string.Empty },
                { "SUPPRESS_UI_ON_OPEN", "0" },
                { "SUPPRESS_VSTO_HOME_ON_ACTIVATE", "0" }
            };

        private static readonly IReadOnlyDictionary<string, string> BaseFixedValues =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "SYSTEM_ROOT", string.Empty },
                { "WORD_TEMPLATE_DIR", string.Empty },
                { "TASKPANE_SNAPSHOT_CACHE_COUNT", "0" }
            };

        private static readonly IReadOnlyList<string> KernelClearPrefixNames = new string[0];

        private static readonly IReadOnlyList<string> BaseClearPrefixNames = new[]
        {
            "TASKPANE_SNAPSHOT_CACHE_"
        };

        /// <summary>
        internal WorkbookResetDefinition GetKernelDefinition()
        {
            return new WorkbookResetDefinition("Kernel", KernelFixedValues, KernelClearPrefixNames);
        }

        /// <summary>
        internal WorkbookResetDefinition GetBaseDefinition()
        {
            return new WorkbookResetDefinition("Base", BaseFixedValues, BaseClearPrefixNames);
        }
    }
}
