using System;
using System.Text;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal static class KernelNamingService
    {
        internal static string NormalizeNameRuleA(string value)
        {
            switch ((value ?? string.Empty).Trim().ToUpperInvariant())
            {
                case "YYYY":
                case "YYYYMMDD":
                case "YYYYMMDD_":
                case "YYYY-MM-DD":
                case "YYYY/MM/DD":
                case "YYYYMM":
                case "YYYYMM_":
                case "YYYY-MM":
                case "YYYY/MM":
                    return "YYYY";
                case "YY":
                case "YYMMDD":
                case "YYMMDD_":
                case "YY-MM-DD":
                case "YY/MM/DD":
                case "YYMM":
                case "YYMM_":
                case "YY-MM":
                case "YY/MM":
                    return "YY";
                case "":
                case "NONE":
                case "OFF":
                case "NO":
                    return "NONE";
                default:
                    return "NONE";
            }
        }

        internal static string NormalizeNameRuleB(string value)
        {
            switch ((value ?? string.Empty).Trim().ToUpperInvariant())
            {
                case "DOC":
                case "DOCUMENT":
                case "DOCONLY":
                    return "DOC";
                case "DOC_CUST":
                case "DOC-CUST":
                case "DOCCUST":
                case "DOC_CUS":
                    return "DOC_CUST";
                case "CUST_DOC":
                case "CUST-DOC":
                case "CUSTDOC":
                    return "CUST_DOC";
                default:
                    return "DOC";
            }
        }

        internal static string BuildDatePrefix(string ruleA, DateTime today)
        {
            switch (NormalizeNameRuleA(ruleA))
            {
                case "YYYY":
                    return today.ToString("yyyyMMdd") + "_";
                case "YY":
                    return today.ToString("yyMMdd") + "_";
                default:
                    return string.Empty;
            }
        }

        internal static string BuildDocumentName(string ruleA, string ruleB, string documentName, string customerName, DateTime today)
        {
            string normalizedCustomer = (customerName ?? string.Empty).Trim();
            string body = documentName ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(normalizedCustomer))
            {
                switch (NormalizeNameRuleB(ruleB))
                {
                    case "DOC_CUST":
                        body += "_" + normalizedCustomer;
                        break;
                    case "CUST_DOC":
                        body = normalizedCustomer + "_" + body;
                        break;
                }
            }

            return BuildDatePrefix(ruleA, today) + body;
        }

        internal static string BuildFolderName(string ruleA, string customerName, DateTime today)
        {
            return BuildDatePrefix(ruleA, today) + SanitizeFolderNameText(customerName);
        }

        internal static string BuildCaseBookName(string customerName, string extension)
        {
            return Infrastructure.WorkbookFileNameResolver.BuildCaseWorkbookName(customerName, extension);
        }

        internal static string SanitizeFileNameForCase(string value)
        {
            return NormalizeNameText(value, " ", false, string.Empty);
        }

        internal static string SanitizeFolderNameText(string value)
        {
            return NormalizeNameText(value, " ", true, string.Empty);
        }

        private static string NormalizeNameText(string value, string badCharReplacement, bool collapseSpaces, string emptyFallback)
        {
            string text = (value ?? string.Empty).Trim();
            string[] badChars = { "\\", "/", ":", "*", "?", "\"", "<", ">", "|" };
            foreach (string badChar in badChars)
            {
                text = text.Replace(badChar, badCharReplacement);
            }

            StringBuilder builder = new StringBuilder(text.Length);
            foreach (char ch in text)
            {
                if (!char.IsControl(ch))
                {
                    builder.Append(ch);
                }
            }

            text = builder.ToString();
            if (collapseSpaces)
            {
                while (text.Contains("  "))
                {
                    text = text.Replace("  ", " ");
                }
            }

            text = text.TrimEnd('.', ' ');
            return text.Length == 0 ? emptyFallback : text;
        }
    }
}
