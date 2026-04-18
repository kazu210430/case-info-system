using System;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal enum DocumentCommandActionRoute
    {
        Document,
        Accounting,
        CaseList,
        Unsupported
    }

    internal static class DocumentCommandActionRoutePolicy
    {
        internal static DocumentCommandActionRoute Decide(string actionKind)
        {
            if (string.Equals(actionKind, "doc", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentCommandActionRoute.Document;
            }

            if (string.Equals(actionKind, "accounting", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentCommandActionRoute.Accounting;
            }

            if (string.Equals(actionKind, "caselist", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentCommandActionRoute.CaseList;
            }

            return DocumentCommandActionRoute.Unsupported;
        }
    }
}
