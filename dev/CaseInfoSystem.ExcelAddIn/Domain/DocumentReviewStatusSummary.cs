namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class DocumentReviewStatusSummary
    {
        /// <summary>
        internal int PassCount { get; set; }

        /// <summary>
        internal int HoldCount { get; set; }

        /// <summary>
        internal int FailCount { get; set; }

        /// <summary>
        internal int OtherCount { get; set; }

        /// <summary>
        internal bool HasAnyReview { get; set; }

        /// <summary>
        internal bool HasConflictingKnownStatuses
        {
            get
            {
                int kinds = 0;
                if (PassCount > 0) { kinds++; }
                if (HoldCount > 0) { kinds++; }
                if (FailCount > 0) { kinds++; }
                return kinds > 1;
            }
        }

        /// <summary>
        internal bool HasDuplicateStatuses
        {
            get
            {
                return PassCount > 1 || HoldCount > 1 || FailCount > 1 || OtherCount > 1;
            }
        }
    }
}
