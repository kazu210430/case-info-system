using System;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class CaseListFieldDefinition
    {
        /// <summary>
        internal string FieldKey { get; set; }

        /// <summary>
        internal string Label { get; set; }

        /// <summary>
        internal string SourceCellAddress { get; set; }

        /// <summary>
        internal string SourceNamedRange { get; set; }

        /// <summary>
        internal string DataType { get; set; }

        /// <summary>
        internal string NormalizeRule { get; set; }
    }
}
