using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class CaseDataSnapshot
    {
        /// <summary>
        internal IReadOnlyDictionary<string, string> Values { get; set; }

        /// <summary>
        internal string CustomerName { get; set; }
    }
}
