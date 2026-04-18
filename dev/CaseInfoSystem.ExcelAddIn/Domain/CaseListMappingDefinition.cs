using System;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class CaseListMappingDefinition
    {
        /// <summary>
        internal string MappingType { get; set; }

        /// <summary>
        internal string SourceFieldKey { get; set; }

        /// <summary>
        internal string TargetHeaderName { get; set; }

        /// <summary>
        internal string DataType { get; set; }

        /// <summary>
        internal string NormalizeRule { get; set; }
    }
}
