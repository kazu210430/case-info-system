using System;
using System.Collections.Generic;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentMasterCatalogDiagnosticsService
    {
        private readonly MasterTemplateCatalogService _masterTemplateCatalogService;
        private readonly DocumentExecutionEligibilityService _documentExecutionEligibilityService;
        private readonly DocumentExecutionPolicyService _documentExecutionPolicyService;
        private readonly DocumentExecutionModeService _documentExecutionModeService;
        private readonly Logger _logger;
        private readonly Dictionary<string, string> _lastSignatureByWorkbook;

        /// <summary>
        internal DocumentMasterCatalogDiagnosticsService(
            MasterTemplateCatalogService masterTemplateCatalogService,
            DocumentExecutionEligibilityService documentExecutionEligibilityService,
            DocumentExecutionPolicyService documentExecutionPolicyService,
            DocumentExecutionModeService documentExecutionModeService,
            Logger logger)
        {
            _masterTemplateCatalogService = masterTemplateCatalogService ?? throw new ArgumentNullException(nameof(masterTemplateCatalogService));
            _documentExecutionEligibilityService = documentExecutionEligibilityService ?? throw new ArgumentNullException(nameof(documentExecutionEligibilityService));
            _documentExecutionPolicyService = documentExecutionPolicyService ?? throw new ArgumentNullException(nameof(documentExecutionPolicyService));
            _documentExecutionModeService = documentExecutionModeService ?? throw new ArgumentNullException(nameof(documentExecutionModeService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _lastSignatureByWorkbook = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        internal void TraceAllMasterTemplates(Excel.Workbook caseWorkbook)
        {
            if (caseWorkbook == null)
            {
                return;
            }

            IReadOnlyList<MasterTemplateRecord> records = _masterTemplateCatalogService.GetAllTemplates(caseWorkbook);
            if (records == null || records.Count == 0)
            {
                return;
            }

            string workbookFullName = caseWorkbook.FullName ?? string.Empty;
            List<string> identities = records
                .Where(record => record != null)
                .Select(BuildEntryIdentity)
                .Where(identity => identity.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(identity => identity, StringComparer.OrdinalIgnoreCase)
                .ToList();
            string signature = workbookFullName + "|" + string.Join(",", identities);
            if (_lastSignatureByWorkbook.TryGetValue(workbookFullName, out string previousSignature)
                && string.Equals(previousSignature, signature, StringComparison.Ordinal))
            {
                return;
            }

            _lastSignatureByWorkbook[workbookFullName] = signature;

            var eligibleEntries = new List<string>();
            var passReviewedEntries = new List<string>();
            var allowlistedEntries = new List<string>();
            var rolloutReadyEntries = new List<string>();
            var blockedByReviewEntries = new List<string>();
            var blockedByAllowlistEntries = new List<string>();
            var groupedFallbackReasons = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var blockedByAllowlistSpecs = new List<DocumentTemplateSpec>();

            foreach (MasterTemplateRecord record in records.Where(item => item != null))
            {
                DocumentExecutionEligibility eligibility = _documentExecutionEligibilityService.Evaluate(caseWorkbook, "doc", record.Key);
                string entryIdentity = BuildEntryIdentity(record);
                if (eligibility.CanExecuteInVsto)
                {
                    if (entryIdentity.Length > 0)
                    {
                        eligibleEntries.Add(entryIdentity);
                    }

                    if (_documentExecutionPolicyService.HasPassedReview(eligibility.TemplateSpec))
                    {
                        passReviewedEntries.Add(entryIdentity);
                    }
                    else
                    {
                        blockedByReviewEntries.Add(FormatReviewWarning(entryIdentity, _documentExecutionPolicyService.GetReviewStatusSummary(eligibility.TemplateSpec)));
                    }

                    if (_documentExecutionPolicyService.IsVstoExecutionAllowed(eligibility.TemplateSpec))
                    {
                        allowlistedEntries.Add(entryIdentity);
                    }
                    else if (eligibility.TemplateSpec != null)
                    {
                        blockedByAllowlistEntries.Add(entryIdentity);
                        blockedByAllowlistSpecs.Add(eligibility.TemplateSpec);
                    }

                    if (_documentExecutionPolicyService.IsRolloutReady(eligibility.TemplateSpec))
                    {
                        rolloutReadyEntries.Add(entryIdentity);
                    }

                    continue;
                }

                string reason = string.IsNullOrWhiteSpace(eligibility.Reason) ? "unknown reason" : eligibility.Reason;
                if (!groupedFallbackReasons.TryGetValue(reason, out List<string> values))
                {
                    values = new List<string>();
                    groupedFallbackReasons.Add(reason, values);
                }

                values.Add(entryIdentity.Length == 0 ? (record.Key ?? string.Empty) : entryIdentity);
            }

            _logger.Info(
                "Document master catalog summary."
                + " workbook=" + workbookFullName
                + ", mode=" + _documentExecutionModeService.GetMode().ToString()
                + ", masterListDocuments=" + identities.Count.ToString()
                + ", vstoEligible=" + eligibleEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                + ", passReviewed=" + passReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                + ", allowlisted=" + allowlistedEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString()
                + ", rolloutReady=" + rolloutReadyEntries.Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString());

            if (eligibleEntries.Count > 0)
            {
                _logger.Info("Document master catalog VSTO candidates. keys=" + string.Join(",", eligibleEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (passReviewedEntries.Count > 0)
            {
                _logger.Info("Document master catalog PASS-reviewed candidates. keys=" + string.Join(",", passReviewedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (allowlistedEntries.Count > 0)
            {
                _logger.Info("Document master catalog allowlisted candidates. keys=" + string.Join(",", allowlistedEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (rolloutReadyEntries.Count > 0)
            {
                _logger.Info("Document master catalog rollout-ready candidates. keys=" + string.Join(",", rolloutReadyEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }

            if (blockedByReviewEntries.Count > 0)
            {
                _logger.Info(
                    "Document master catalog candidates blocked by review status. keys="
                    + string.Join(",", blockedByReviewEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                    + ", reviewNotesPath=" + _documentExecutionPolicyService.GetReviewNotesPath());
            }

            if (blockedByAllowlistEntries.Count > 0)
            {
                _logger.Info(
                    "Document master catalog candidates blocked by allowlist. keys="
                    + string.Join(",", blockedByAllowlistEntries.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                    + ", allowlistPath=" + _documentExecutionPolicyService.GetAllowlistPath());
                _logger.Info("Document master catalog allowlist suggestions begin");
                foreach (DocumentTemplateSpec templateSpec in blockedByAllowlistSpecs
                    .Where(spec => spec != null)
                    .GroupBy(BuildEntryIdentity, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .OrderBy(BuildEntryIdentity, StringComparer.OrdinalIgnoreCase))
                {
                    _logger.Info(BuildAllowlistFileSuggestion(templateSpec));
                    _logger.Info(BuildReviewNotesSuggestion(templateSpec));
                }
                _logger.Info("Document master catalog allowlist suggestions end");
            }

            foreach (KeyValuePair<string, List<string>> pair in groupedFallbackReasons.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                _logger.Info("Document master catalog non-VSTO reasons. reason=" + pair.Key + ", keys=" + string.Join(",", pair.Value.OrderBy(value => value, StringComparer.OrdinalIgnoreCase)));
            }
        }

        /// <summary>
        private static string BuildEntryIdentity(MasterTemplateRecord record)
        {
            if (record == null)
            {
                return string.Empty;
            }

            string key = (record.Key ?? string.Empty).Trim();
            string templateFileName = (record.TemplateFileName ?? string.Empty).Trim();
            if (key.Length == 0 || templateFileName.Length == 0)
            {
                return string.Empty;
            }

            return key + "|" + templateFileName;
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
        private static string FormatReviewWarning(string entryIdentity, DocumentReviewStatusSummary summary)
        {
            DocumentReviewStatusSummary safeSummary = summary ?? new DocumentReviewStatusSummary();
            return (entryIdentity ?? string.Empty)
                + "(PASS=" + safeSummary.PassCount.ToString()
                + ",HOLD=" + safeSummary.HoldCount.ToString()
                + ",FAIL=" + safeSummary.FailCount.ToString()
                + ",OTHER=" + safeSummary.OtherCount.ToString()
                + ")";
        }

        /// <summary>
        private static string BuildAllowlistFileSuggestion(DocumentTemplateSpec templateSpec)
        {
            string identity = BuildEntryIdentity(templateSpec);
            return identity.Length == 0 ? string.Empty : "ALLOWLIST_FILE_CANDIDATE " + identity;
        }

        /// <summary>
        private static string BuildReviewNotesSuggestion(DocumentTemplateSpec templateSpec)
        {
            string identity = BuildEntryIdentity(templateSpec);
            return identity.Length == 0 ? string.Empty : "ALLOWLIST_REVIEW_CANDIDATE " + identity + "|HOLD|yyyy-MM-dd|reviewer|notes";
        }
    }
}
