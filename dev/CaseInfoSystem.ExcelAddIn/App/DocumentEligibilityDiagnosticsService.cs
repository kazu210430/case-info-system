using System;
using System.Collections.Generic;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentEligibilityDiagnosticsService
    {
        private readonly DocumentExecutionModeService _documentExecutionModeService;
        private readonly DocumentExecutionEligibilityService _documentExecutionEligibilityService;
        private readonly DocumentExecutionPolicyService _documentExecutionPolicyService;
        private readonly Logger _logger;
        private readonly Dictionary<string, string> _lastSignatureByWorkbook;

        /// <summary>
        internal DocumentEligibilityDiagnosticsService(
            DocumentExecutionModeService documentExecutionModeService,
            DocumentExecutionEligibilityService documentExecutionEligibilityService,
            DocumentExecutionPolicyService documentExecutionPolicyService,
            Logger logger)
        {
            _documentExecutionModeService = documentExecutionModeService ?? throw new ArgumentNullException(nameof(documentExecutionModeService));
            _documentExecutionEligibilityService = documentExecutionEligibilityService ?? throw new ArgumentNullException(nameof(documentExecutionEligibilityService));
            _documentExecutionPolicyService = documentExecutionPolicyService ?? throw new ArgumentNullException(nameof(documentExecutionPolicyService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _lastSignatureByWorkbook = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        internal void TraceCaseSnapshot(Excel.Workbook workbook, TaskPaneSnapshot snapshot)
        {
            if (workbook == null || snapshot == null || snapshot.HasError)
            {
                return;
            }

            List<string> keys = snapshot.DocButtons
                .Where(doc => string.Equals(doc.ActionKind, "doc", StringComparison.OrdinalIgnoreCase))
                .Select(doc => (doc.Key ?? string.Empty).Trim())
                .Where(key => key.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (keys.Count == 0)
            {
                return;
            }

            string workbookFullName = workbook.FullName ?? string.Empty;
            string signature = workbookFullName + "|" + snapshot.MasterVersion.ToString() + "|" + string.Join(",", keys);
            if (_lastSignatureByWorkbook.TryGetValue(workbookFullName, out string previousSignature)
                && string.Equals(previousSignature, signature, StringComparison.Ordinal))
            {
                return;
            }

            var eligibleKeys = new List<string>();
            var allowlistedKeys = new List<string>();
            var allowlistedWithoutPassedReview = new List<string>();
            var passReviewedEligibleKeys = new List<string>();
            var passReviewedAllowlistedKeys = new List<string>();
            var passReviewedBlockedByAllowlist = new List<string>();
            var eligibleWithoutPassedReview = new List<string>();
            var resolvedTemplateDetails = new List<string>();
            var resolutionSourceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var blockedByPolicySpecs = new List<DocumentTemplateSpec>();
            var currentResolvedEntryIdentities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var groupedFallbackReasons = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var currentPassReviewedEntries = new List<string>();
            var currentHoldReviewedEntries = new List<string>();
            var currentFailReviewedEntries = new List<string>();
            var currentOtherReviewedEntries = new List<string>();
            var currentUnreviewedEntries = new List<string>();

            foreach (string key in keys)
            {
                DocumentExecutionEligibility eligibility = _documentExecutionEligibilityService.Evaluate(workbook, "doc", key);
                string currentEntryIdentity = BuildEntryIdentity(eligibility.TemplateSpec);
                if (currentEntryIdentity.Length > 0)
                {
                    currentResolvedEntryIdentities.Add(currentEntryIdentity);
                    resolvedTemplateDetails.Add(BuildResolvedTemplateDetail(eligibility));

                    string resolutionSourceKey = eligibility.ResolutionSource.ToString();
                    if (!resolutionSourceCounts.ContainsKey(resolutionSourceKey))
                    {
                        resolutionSourceCounts[resolutionSourceKey] = 0;
                    }

                    resolutionSourceCounts[resolutionSourceKey]++;

                    DocumentReviewStatusSummary currentReviewStatus = _documentExecutionPolicyService.GetReviewStatusSummary(currentEntryIdentity);
                    if (currentReviewStatus.PassCount > 0)
                    {
                        currentPassReviewedEntries.Add(currentEntryIdentity);
                    }
                    else if (currentReviewStatus.HoldCount > 0)
                    {
                        currentHoldReviewedEntries.Add(FormatReviewWarning(currentEntryIdentity, currentReviewStatus));
                    }
                    else if (currentReviewStatus.FailCount > 0)
                    {
                        currentFailReviewedEntries.Add(FormatReviewWarning(currentEntryIdentity, currentReviewStatus));
                    }
                    else if (currentReviewStatus.OtherCount > 0)
                    {
                        currentOtherReviewedEntries.Add(FormatReviewWarning(currentEntryIdentity, currentReviewStatus));
                    }
                    else
                    {
                        currentUnreviewedEntries.Add(currentEntryIdentity);
                    }
                }

                if (eligibility.CanExecuteInVsto)
                {
                    eligibleKeys.Add(key);
                    string formattedTemplateKey = FormatTemplateKey(eligibility.TemplateSpec);
                    DocumentReviewStatusSummary reviewStatusSummary = _documentExecutionPolicyService.GetReviewStatusSummary(eligibility.TemplateSpec);
                    if (_documentExecutionPolicyService.HasPassedReview(eligibility.TemplateSpec))
                    {
                        if (formattedTemplateKey.Length > 0)
                        {
                            passReviewedEligibleKeys.Add(formattedTemplateKey);
                        }
                    }
                    else if (formattedTemplateKey.Length > 0)
                    {
                        eligibleWithoutPassedReview.Add(FormatReviewWarning(formattedTemplateKey, reviewStatusSummary));
                    }

                    if (_documentExecutionPolicyService.IsVstoExecutionAllowed(eligibility.TemplateSpec))
                    {
                        allowlistedKeys.Add(formattedTemplateKey);
                        if (_documentExecutionPolicyService.HasPassedReview(eligibility.TemplateSpec))
                        {
                            passReviewedAllowlistedKeys.Add(formattedTemplateKey);
                        }
                        else
                        {
                            allowlistedWithoutPassedReview.Add(FormatReviewWarning(formattedTemplateKey, reviewStatusSummary));
                        }
                    }
                    else if (eligibility.TemplateSpec != null)
                    {
                        blockedByPolicySpecs.Add(eligibility.TemplateSpec);
                        if (_documentExecutionPolicyService.HasPassedReview(eligibility.TemplateSpec) && formattedTemplateKey.Length > 0)
                        {
                            passReviewedBlockedByAllowlist.Add(formattedTemplateKey);
                        }
                    }
                    continue;
                }

                string reason = string.IsNullOrWhiteSpace(eligibility.Reason) ? "unknown reason" : eligibility.Reason;
                if (!groupedFallbackReasons.TryGetValue(reason, out List<string> values))
                {
                    values = new List<string>();
                    groupedFallbackReasons.Add(reason, values);
                }

                values.Add(key);
            }

            _lastSignatureByWorkbook[workbookFullName] = signature;

            _logger.Info(
                "Document eligibility summary."
                + " workbook=" + workbookFullName
                + ", masterVersion=" + snapshot.MasterVersion.ToString()
                + ", mode=" + _documentExecutionModeService.GetMode().ToString()
                + ", docButtons=" + keys.Count.ToString()
                + ", resolvedEntries=" + currentResolvedEntryIdentities.Count.ToString()
                + ", vstoEligible=" + eligibleKeys.Count.ToString()
                + ", allowlisted=" + allowlistedKeys.Count.ToString()
                + ", nonVsto=" + groupedFallbackReasons.Values.Sum(list => list.Count).ToString());

            if (currentResolvedEntryIdentities.Count > 0)
            {
                _logger.Info(
                    "Document eligibility resolved template identities. keys="
                    + string.Join(",", currentResolvedEntryIdentities.OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (resolutionSourceCounts.Count > 0)
            {
                _logger.Info(
                    "Document eligibility resolution sources. counts="
                    + string.Join(",", resolutionSourceCounts
                        .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(pair => pair.Key + "=" + pair.Value.ToString())));
            }

            if (resolvedTemplateDetails.Count > 0)
            {
                _logger.Info(
                    "Document eligibility resolved template details. values="
                    + string.Join(",", resolvedTemplateDetails
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (currentResolvedEntryIdentities.Count > 0)
            {
                _logger.Info(
                    "Document eligibility current review summary."
                    + " pass=" + currentPassReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                    + ", hold=" + currentHoldReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                    + ", fail=" + currentFailReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                    + ", other=" + currentOtherReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                    + ", unreviewed=" + currentUnreviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            if (currentPassReviewedEntries.Count > 0)
            {
                _logger.Info(
                    "Document eligibility current PASS-reviewed entries. keys="
                    + string.Join(",", currentPassReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (currentHoldReviewedEntries.Count > 0)
            {
                _logger.Info(
                    "Document eligibility current HOLD-reviewed entries. keys="
                    + string.Join(",", currentHoldReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (currentFailReviewedEntries.Count > 0)
            {
                _logger.Warn(
                    "Document eligibility current FAIL-reviewed entries."
                    + " keys=" + string.Join(",", currentFailReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (currentOtherReviewedEntries.Count > 0)
            {
                _logger.Warn(
                    "Document eligibility current OTHER-reviewed entries."
                    + " keys=" + string.Join(",", currentOtherReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (currentUnreviewedEntries.Count > 0)
            {
                _logger.Info(
                    "Document eligibility current unreviewed entries. keys="
                    + string.Join(",", currentUnreviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            List<string> conflictingReviewedEntries = _documentExecutionPolicyService
                .GetConflictingReviewedDocumentIdentities()
                .Where(identity => currentResolvedEntryIdentities.Contains(identity))
                .Select(identity => FormatReviewWarning(identity, _documentExecutionPolicyService.GetReviewStatusSummary(identity)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (conflictingReviewedEntries.Count > 0)
            {
                _logger.Warn(
                    "Document eligibility current review entries contain conflicting statuses."
                    + " keys=" + string.Join(",", conflictingReviewedEntries)
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            List<string> duplicateReviewedEntries = _documentExecutionPolicyService
                .GetDuplicateReviewedDocumentIdentities()
                .Where(identity => currentResolvedEntryIdentities.Contains(identity))
                .Select(identity => FormatReviewWarning(identity, _documentExecutionPolicyService.GetReviewStatusSummary(identity)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (duplicateReviewedEntries.Count > 0)
            {
                _logger.Warn(
                    "Document eligibility current review entries contain duplicate statuses."
                    + " keys=" + string.Join(",", duplicateReviewedEntries)
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            if (eligibleKeys.Count > 0)
            {
                _logger.Info("Document eligibility VSTO candidates. keys=" + string.Join(",", eligibleKeys));
            }

            if (passReviewedEligibleKeys.Count > 0)
            {
                _logger.Info(
                    "Document eligibility PASS-reviewed VSTO candidates. keys="
                    + string.Join(",", passReviewedEligibleKeys.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (passReviewedAllowlistedKeys.Count > 0)
            {
                _logger.Info(
                    "Document eligibility rollout-ready allowlisted candidates. keys="
                    + string.Join(",", passReviewedAllowlistedKeys.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (passReviewedBlockedByAllowlist.Count > 0)
            {
                _logger.Info(
                    "Document eligibility PASS-reviewed candidates blocked only by allowlist. keys="
                    + string.Join(",", passReviewedBlockedByAllowlist.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                    + ", allowlistPath=" + _documentExecutionPolicyService.GetAllowlistPath());
                _logger.Info("Document eligibility rollout-ready allowlist suggestions begin");
                foreach (DocumentTemplateSpec templateSpec in blockedByPolicySpecs
                    .Where(spec => spec != null && _documentExecutionPolicyService.HasPassedReview(spec))
                    .GroupBy(spec => FormatTemplateKey(spec), StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .OrderBy(spec => FormatTemplateKey(spec), StringComparer.OrdinalIgnoreCase))
                {
                    _logger.Info(BuildAllowlistSuggestion(templateSpec));
                    _logger.Info(BuildAllowlistFileSuggestion(templateSpec));
                }
                _logger.Info("Document eligibility rollout-ready allowlist suggestions end");
            }

            if (eligibleWithoutPassedReview.Count > 0)
            {
                _logger.Info(
                    "Document eligibility candidates still blocked by review status. keys="
                    + string.Join(",", eligibleWithoutPassedReview.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            if (allowlistedKeys.Count > 0)
            {
                _logger.Info("Document eligibility VSTO allowlisted. keys=" + string.Join(",", allowlistedKeys));
            }

            if (allowlistedWithoutPassedReview.Count > 0)
            {
                _logger.Warn(
                    "Document eligibility allowlisted documents without PASS review notes."
                    + " keys=" + string.Join(",", allowlistedWithoutPassedReview.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            List<string> staleAllowlistedEntries = _documentExecutionPolicyService
                .GetAllowedDocuments()
                .Select(BuildEntryIdentity)
                .Where(identity => identity.Length > 0 && !currentResolvedEntryIdentities.Contains(identity))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(identity => identity, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (staleAllowlistedEntries.Count > 0)
            {
                _logger.Warn(
                    "Document eligibility allowlist entries not present in current snapshot."
                    + " keys=" + string.Join(",", staleAllowlistedEntries)
                    + ", allowlistPath=" + _documentExecutionPolicyService.GetAllowlistPath());
            }

            List<string> staleReviewedEntries = _documentExecutionPolicyService
                .GetReviewedDocumentIdentities()
                .Where(identity => !currentResolvedEntryIdentities.Contains(identity))
                .Select(identity => FormatReviewWarning(identity, _documentExecutionPolicyService.GetReviewStatusSummary(identity)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(identity => identity, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (staleReviewedEntries.Count > 0)
            {
                _logger.Info(
                    "Document eligibility review entries not present in current snapshot."
                    + " keys=" + string.Join(",", staleReviewedEntries)
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            if (staleAllowlistedEntries.Count > 0 || staleReviewedEntries.Count > 0)
            {
                _logger.Info(
                    "Document eligibility stale policy summary."
                    + " staleAllowlistEntries=" + staleAllowlistedEntries.Count.ToString()
                    + ", staleReviewEntries=" + staleReviewedEntries.Count.ToString());
            }

            if (eligibleKeys.Count > 0 && allowlistedKeys.Count != eligibleKeys.Count)
            {
                List<string> blockedByPolicy = blockedByPolicySpecs
                    .Select(FormatTemplateKey)
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                _logger.Info("Document eligibility blocked by allowlist. keys=" + string.Join(",", blockedByPolicy));
                _logger.Info("Document eligibility policy files. allowlistPath=" + _documentExecutionPolicyService.GetAllowlistPath() + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
                _logger.Info("Document eligibility allowlist suggestions begin");
                foreach (DocumentTemplateSpec templateSpec in blockedByPolicySpecs
                    .Where(spec => spec != null)
                    .GroupBy(spec => FormatTemplateKey(spec), StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .OrderBy(spec => FormatTemplateKey(spec), StringComparer.OrdinalIgnoreCase))
                {
                    _logger.Info(BuildAllowlistSuggestion(templateSpec));
                    _logger.Info(BuildAllowlistFileSuggestion(templateSpec));
                    _logger.Info(BuildReviewNotesSuggestion(templateSpec));
                }
                _logger.Info("Document eligibility allowlist suggestions end");
            }

            foreach (KeyValuePair<string, List<string>> pair in groupedFallbackReasons.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                _logger.Info("Document eligibility non-VSTO reasons. reason=" + pair.Key + ", keys=" + string.Join(",", pair.Value));
            }
        }

        /// <summary>
        private static string FormatTemplateKey(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null)
            {
                return string.Empty;
            }

            string key = (templateSpec.Key ?? string.Empty).Trim();
            string templateFileName = (templateSpec.TemplateFileName ?? string.Empty).Trim();
            return key + ":" + templateFileName;
        }

        /// <summary>
        private static string BuildEntryIdentity(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null)
            {
                return string.Empty;
            }

            string key = (templateSpec.Key ?? string.Empty).Trim();
            string templateFileName = (templateSpec.TemplateFileName ?? string.Empty).Trim();
            if (key.Length == 0 || templateFileName.Length == 0)
            {
                return string.Empty;
            }

            return key + "|" + templateFileName;
        }

        /// <summary>
        private static string BuildEntryIdentity(DocumentExecutionPolicyEntry policyEntry)
        {
            if (policyEntry == null)
            {
                return string.Empty;
            }

            string key = (policyEntry.Key ?? string.Empty).Trim();
            string templateFileName = (policyEntry.TemplateFileName ?? string.Empty).Trim();
            if (key.Length == 0 || templateFileName.Length == 0)
            {
                return string.Empty;
            }

            return key + "|" + templateFileName;
        }

        /// <summary>
        private static string BuildResolvedTemplateDetail(DocumentExecutionEligibility eligibility)
        {
            if (eligibility == null || eligibility.TemplateSpec == null)
            {
                return string.Empty;
            }

            DocumentTemplateSpec templateSpec = eligibility.TemplateSpec;
            string key = (templateSpec.Key ?? string.Empty).Trim();
            string templateFileName = (templateSpec.TemplateFileName ?? string.Empty).Trim();
            string documentName = (templateSpec.DocumentName ?? string.Empty).Trim();
            string resolutionSource = eligibility.ResolutionSource.ToString();
            return key + ":" + templateFileName + "[" + resolutionSource + "|" + documentName + "]";
        }

        /// <summary>
        private static string BuildAllowlistSuggestion(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null)
            {
                return string.Empty;
            }

            string key = EscapeCSharpString((templateSpec.Key ?? string.Empty).Trim());
            string templateFileName = EscapeCSharpString((templateSpec.TemplateFileName ?? string.Empty).Trim());
            return "ALLOWLIST_CANDIDATE new DocumentExecutionPolicyEntry { Key = \"" + key + "\", TemplateFileName = \"" + templateFileName + "\" },";
        }

        /// <summary>
        private static string BuildAllowlistFileSuggestion(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null)
            {
                return string.Empty;
            }

            string key = (templateSpec.Key ?? string.Empty).Trim();
            string templateFileName = (templateSpec.TemplateFileName ?? string.Empty).Trim();
            return "ALLOWLIST_FILE_CANDIDATE " + key + "|" + templateFileName;
        }

        /// <summary>
        private static string BuildReviewNotesSuggestion(DocumentTemplateSpec templateSpec)
        {
            if (templateSpec == null)
            {
                return string.Empty;
            }

            string key = (templateSpec.Key ?? string.Empty).Trim();
            string templateFileName = (templateSpec.TemplateFileName ?? string.Empty).Trim();
            return "ALLOWLIST_REVIEW_CANDIDATE " + key + "|" + templateFileName + "|HOLD|yyyy-MM-dd|reviewer|notes";
        }

        /// <summary>
        private static string EscapeCSharpString(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        /// <summary>
        private static string FormatReviewWarning(string formattedTemplateKey, DocumentReviewStatusSummary reviewStatusSummary)
        {
            DocumentReviewStatusSummary summary = reviewStatusSummary ?? new DocumentReviewStatusSummary();
            return formattedTemplateKey
                + "(PASS=" + summary.PassCount.ToString()
                + ",HOLD=" + summary.HoldCount.ToString()
                + ",FAIL=" + summary.FailCount.ToString()
                + ",OTHER=" + summary.OtherCount.ToString()
                + ")";
        }
    }
}
