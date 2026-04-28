using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class DocumentExecutionPolicyService
	{
		private sealed class ReviewEntryInfo
		{
			internal string EntryIdentity { get; set; } = string.Empty;

			internal string Status { get; set; } = string.Empty;

			internal string ReviewedOn { get; set; } = string.Empty;

			internal string Reviewer { get; set; } = string.Empty;

			internal string Notes { get; set; } = string.Empty;
		}

		private const char EntrySeparator = '|';

		private const string AllowlistFileName = "DocumentExecutionAllowlist.txt";

		private const string ReviewNotesFileName = "DocumentExecutionAllowlist.review.txt";

		private const int ReviewNotesMinimumColumnCount = 6;

		private const string ReviewStatusPass = "PASS";

		private const string ReviewStatusHold = "HOLD";

		private const string ReviewStatusFail = "FAIL";

		private const string ReviewDateFormat = "yyyy-MM-dd";

		private const string DefaultSystemRootFolderName = "案件情報System";

		private const string AddInsFolderName = "Addins";

		private const string RuntimeAddInFolderName = "CaseInfoSystem.ExcelAddIn";

		private const string SystemRootPropertyName = "SYSTEM_ROOT";

		private static readonly IReadOnlyList<DocumentExecutionPolicyEntry> BuiltInAllowedDocuments = new DocumentExecutionPolicyEntry[0];

		private readonly Logger _logger;

		private readonly ExcelInteropService _excelInteropService;

		private IReadOnlyList<DocumentExecutionPolicyEntry> _allowedDocuments;

		private HashSet<string> _reviewPassedKeys;

		private Dictionary<string, DocumentReviewStatusSummary> _reviewStatusByEntryIdentity;

		private string _loadedAllowlistPath;

		private string _loadedReviewNotesPath;

		private long _loadedAllowlistLastWriteUtcTicks;

		private long _loadedReviewNotesLastWriteUtcTicks;

		internal DocumentExecutionPolicyService (Logger logger, ExcelInteropService excelInteropService)
		{
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_allowedDocuments = Array.Empty<DocumentExecutionPolicyEntry> ();
			_reviewPassedKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
			_reviewStatusByEntryIdentity = new Dictionary<string, DocumentReviewStatusSummary> (StringComparer.OrdinalIgnoreCase);
			_loadedAllowlistPath = string.Empty;
			_loadedReviewNotesPath = string.Empty;
			_loadedAllowlistLastWriteUtcTicks = long.MinValue;
			_loadedReviewNotesLastWriteUtcTicks = long.MinValue;
		}

		// These compatibility APIs no longer gate document execution at runtime.
		// They remain permissive so diagnostics and legacy tooling can still read policy files safely.
		internal bool IsVstoExecutionAllowed (DocumentTemplateSpec templateSpec)
		{
			return templateSpec != null;
		}

		internal IReadOnlyCollection<DocumentExecutionPolicyEntry> GetAllowedDocuments ()
		{
			EnsureLoaded ();
			return _allowedDocuments;
		}

		internal IReadOnlyCollection<string> GetReviewedDocumentIdentities ()
		{
			EnsureLoaded ();
			return (IReadOnlyCollection<string>)(object)_reviewStatusByEntryIdentity.Keys.ToArray ();
		}

		internal string GetAllowlistPath ()
		{
			return ResolvePolicyFilePath ("DocumentExecutionAllowlist.txt");
		}

		internal string GetReviewNotesPath ()
		{
			return ResolvePolicyFilePath ("DocumentExecutionAllowlist.review.txt");
		}

		internal bool HasPassedReview (DocumentTemplateSpec templateSpec)
		{
			return templateSpec != null;
		}

		internal bool IsRolloutReady (DocumentTemplateSpec templateSpec)
		{
			return templateSpec != null;
		}

		internal DocumentReviewStatusSummary GetReviewStatusSummary (DocumentTemplateSpec templateSpec)
		{
			EnsureLoaded ();
			if (templateSpec == null) {
				return new DocumentReviewStatusSummary ();
			}
			string key = BuildEntryIdentity (templateSpec.Key, templateSpec.TemplateFileName);
			if (_reviewStatusByEntryIdentity.TryGetValue (key, out var value) && value != null) {
				return value;
			}
			return new DocumentReviewStatusSummary ();
		}

		internal DocumentReviewStatusSummary GetReviewStatusSummary (string entryIdentity)
		{
			EnsureLoaded ();
			string text = (entryIdentity ?? string.Empty).Trim ();
			if (text.Length == 0) {
				return new DocumentReviewStatusSummary ();
			}
			if (_reviewStatusByEntryIdentity.TryGetValue (text, out var value) && value != null) {
				return value;
			}
			return new DocumentReviewStatusSummary ();
		}

		internal IReadOnlyCollection<string> GetConflictingReviewedDocumentIdentities ()
		{
			EnsureLoaded ();
			return (IReadOnlyCollection<string>)(object)(from pair in _reviewStatusByEntryIdentity
				where pair.Value != null && pair.Value.HasConflictingKnownStatuses
				select pair.Key).OrderBy ((string value) => value, StringComparer.OrdinalIgnoreCase).ToArray ();
		}

		internal IReadOnlyCollection<string> GetDuplicateReviewedDocumentIdentities ()
		{
			EnsureLoaded ();
			return (IReadOnlyCollection<string>)(object)(from pair in _reviewStatusByEntryIdentity
				where pair.Value != null && pair.Value.HasDuplicateStatuses
				select pair.Key).OrderBy ((string value) => value, StringComparer.OrdinalIgnoreCase).ToArray ();
		}

		private IReadOnlyList<DocumentExecutionPolicyEntry> LoadAllowedDocuments ()
		{
			List<DocumentExecutionPolicyEntry> list = new List<DocumentExecutionPolicyEntry> ();
			foreach (DocumentExecutionPolicyEntry builtInAllowedDocument in BuiltInAllowedDocuments) {
				AddEntryIfValid (list, builtInAllowedDocument);
			}
			foreach (DocumentExecutionPolicyEntry item in LoadEntriesFromFile ("DocumentExecutionAllowlist.txt", "Document execution allowlist file")) {
				AddEntryIfValid (list, item);
			}
			return list;
		}

		private Dictionary<string, DocumentReviewStatusSummary> LoadReviewStatusByEntryIdentity (HashSet<string> passedKeys)
		{
			string reviewNotesPath = GetReviewNotesPath ();
			Dictionary<string, DocumentReviewStatusSummary> dictionary = new Dictionary<string, DocumentReviewStatusSummary> (StringComparer.OrdinalIgnoreCase);
			if (!File.Exists (reviewNotesPath)) {
				_logger.Info ("Document execution review notes file was not found. path=" + reviewNotesPath + ", passEntries=0");
				return dictionary;
			}
			try {
				string[] array = File.ReadAllLines (reviewNotesPath);
				for (int i = 0; i < array.Length; i++) {
					string rawLine = array [i];
					ReviewEntryInfo reviewEntryInfo = ParseReviewEntry (rawLine, i + 1, reviewNotesPath);
					if (reviewEntryInfo != null && !string.IsNullOrWhiteSpace (reviewEntryInfo.EntryIdentity)) {
						if (!dictionary.TryGetValue (reviewEntryInfo.EntryIdentity, out var value) || value == null) {
							value = new DocumentReviewStatusSummary ();
							dictionary [reviewEntryInfo.EntryIdentity] = value;
						}
						value.HasAnyReview = true;
						ValidateReviewEntryMetadata (reviewEntryInfo, i + 1, reviewNotesPath);
						if (string.Equals (reviewEntryInfo.Status, "PASS", StringComparison.OrdinalIgnoreCase)) {
							value.PassCount++;
							passedKeys?.Add (reviewEntryInfo.EntryIdentity);
						} else if (string.Equals (reviewEntryInfo.Status, "HOLD", StringComparison.OrdinalIgnoreCase)) {
							value.HoldCount++;
						} else if (string.Equals (reviewEntryInfo.Status, "FAIL", StringComparison.OrdinalIgnoreCase)) {
							value.FailCount++;
						} else {
							value.OtherCount++;
						}
					}
				}
				LogReviewStatusConsistency (dictionary, reviewNotesPath);
				_logger.Info ("Document execution review notes loaded. path=" + reviewNotesPath + ", passEntries=" + passedKeys.Count);
				return dictionary;
			} catch (Exception exception) {
				_logger.Error ("Document execution review notes load failed. path=" + reviewNotesPath, exception);
				return dictionary;
			}
		}

		private string ResolvePolicyFilePath (string fileName)
		{
			string path = fileName ?? string.Empty;
			string text = ResolveRuntimePolicyDirectory ();
			if (!string.IsNullOrWhiteSpace (text)) {
				return Path.Combine (text, path);
			}
			string path2 = Path.GetDirectoryName (typeof(DocumentExecutionPolicyService).Assembly.Location) ?? string.Empty;
			return Path.Combine (path2, path);
		}

		private void EnsureLoaded ()
		{
			string text = ResolvePolicyFilePath ("DocumentExecutionAllowlist.txt");
			string text2 = ResolvePolicyFilePath ("DocumentExecutionAllowlist.review.txt");
			long fileLastWriteUtcTicks = GetFileLastWriteUtcTicks (text);
			long fileLastWriteUtcTicks2 = GetFileLastWriteUtcTicks (text2);
			if (!string.Equals (_loadedAllowlistPath, text, StringComparison.OrdinalIgnoreCase) || !string.Equals (_loadedReviewNotesPath, text2, StringComparison.OrdinalIgnoreCase) || _loadedAllowlistLastWriteUtcTicks != fileLastWriteUtcTicks || _loadedReviewNotesLastWriteUtcTicks != fileLastWriteUtcTicks2) {
				_loadedAllowlistPath = text;
				_loadedReviewNotesPath = text2;
				_loadedAllowlistLastWriteUtcTicks = fileLastWriteUtcTicks;
				_loadedReviewNotesLastWriteUtcTicks = fileLastWriteUtcTicks2;
				_allowedDocuments = LoadAllowedDocuments ();
				_reviewPassedKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
				_reviewStatusByEntryIdentity = LoadReviewStatusByEntryIdentity (_reviewPassedKeys);
				_logger.Info ("Document execution policy cache refreshed. allowlistPath=" + text + ", allowlistLastWriteUtcTicks=" + fileLastWriteUtcTicks.ToString (CultureInfo.InvariantCulture) + ", reviewNotesPath=" + text2 + ", reviewNotesLastWriteUtcTicks=" + fileLastWriteUtcTicks2.ToString (CultureInfo.InvariantCulture));
			}
		}

		private static long GetFileLastWriteUtcTicks (string filePath)
		{
			string text = (filePath ?? string.Empty).Trim ();
			if (text.Length == 0 || !File.Exists (text)) {
				return long.MinValue;
			}
			return File.GetLastWriteTimeUtc (text).Ticks;
		}

		private string ResolveRuntimePolicyDirectory ()
		{
			string text = ResolveRuntimePolicyDirectoryFromOpenWorkbooks ();
			if (!string.IsNullOrWhiteSpace (text)) {
				return text;
			}
			string folderPath = Environment.GetFolderPath (Environment.SpecialFolder.Personal);
			if (string.IsNullOrWhiteSpace (folderPath)) {
				string text2 = Path.GetDirectoryName (typeof(DocumentExecutionPolicyService).Assembly.Location) ?? string.Empty;
				return Directory.Exists (text2) ? text2 : string.Empty;
			}
			string text3 = Path.Combine (folderPath, "案件情報System", "Addins", "CaseInfoSystem.ExcelAddIn");
			if (Directory.Exists (text3)) {
				return text3;
			}
			string text4 = Path.GetDirectoryName (typeof(DocumentExecutionPolicyService).Assembly.Location) ?? string.Empty;
			return Directory.Exists (text4) ? text4 : string.Empty;
		}

		private string ResolveRuntimePolicyDirectoryFromOpenWorkbooks ()
		{
			foreach (Workbook openWorkbook in _excelInteropService.GetOpenWorkbooks ()) {
				string text = (_excelInteropService.TryGetDocumentProperty (openWorkbook, "SYSTEM_ROOT") ?? string.Empty).Trim ();
				if (text.Length != 0) {
					string text2 = Path.Combine (text, "Addins", "CaseInfoSystem.ExcelAddIn");
					if (Directory.Exists (text2)) {
						return text2;
					}
				}
			}
			return string.Empty;
		}

		private IReadOnlyList<DocumentExecutionPolicyEntry> LoadEntriesFromFile (string fileName, string logLabel)
		{
			string text = ResolvePolicyFilePath (fileName);
			List<DocumentExecutionPolicyEntry> list = new List<DocumentExecutionPolicyEntry> ();
			if (!File.Exists (text)) {
				_logger.Info (logLabel + " was not found. path=" + text + ", entries=0");
				return list;
			}
			try {
				string[] array = File.ReadAllLines (text);
				for (int i = 0; i < array.Length; i++) {
					string rawLine = array [i];
					DocumentExecutionPolicyEntry entry = ParseEntry (rawLine, i + 1, text);
					AddEntryIfValid (list, entry);
				}
				_logger.Info (logLabel + " loaded. path=" + text + ", entries=" + list.Count);
				return list;
			} catch (Exception exception) {
				_logger.Error (logLabel + " load failed. path=" + text, exception);
				return list;
			}
		}

		private DocumentExecutionPolicyEntry ParseEntry (string rawLine, int lineNumber, string allowlistPath)
		{
			string text = (rawLine ?? string.Empty).Trim ();
			if (text.Length == 0 || text.StartsWith ("#", StringComparison.Ordinal)) {
				return null;
			}
			string[] array = text.Split (new char[1] { '|' }, 2, StringSplitOptions.None);
			if (array.Length != 2) {
				_logger.Warn ("Document execution allowlist line was ignored. path=" + allowlistPath + ", line=" + lineNumber + ", reason=separator not found");
				return null;
			}
			string text2 = array [0].Trim ();
			string text3 = array [1].Trim ();
			if (text2.Length == 0 || text3.Length == 0) {
				_logger.Warn ("Document execution allowlist line was ignored. path=" + allowlistPath + ", line=" + lineNumber + ", reason=key or templateFileName is empty");
				return null;
			}
			return new DocumentExecutionPolicyEntry {
				Key = text2,
				TemplateFileName = text3
			};
		}

		private ReviewEntryInfo ParseReviewEntry (string rawLine, int lineNumber, string reviewNotesPath)
		{
			string text = (rawLine ?? string.Empty).Trim ();
			if (text.Length == 0 || text.StartsWith ("#", StringComparison.Ordinal)) {
				return null;
			}
			string[] array = text.Split (new char[1] { '|' }, 6, StringSplitOptions.None);
			if (array.Length < 6) {
				_logger.Warn ("Document execution review notes line was ignored. path=" + reviewNotesPath + ", line=" + lineNumber + ", reason=column count is insufficient");
				return null;
			}
			string text2 = array [0].Trim ();
			string text3 = array [1].Trim ();
			string status = array [2].Trim ();
			string reviewedOn = array [3].Trim ();
			string reviewer = array [4].Trim ();
			string notes = array [5].Trim ();
			if (text2.Length == 0 || text3.Length == 0) {
				_logger.Warn ("Document execution review notes line was ignored. path=" + reviewNotesPath + ", line=" + lineNumber + ", reason=key or templateFileName is empty");
				return null;
			}
			return new ReviewEntryInfo {
				EntryIdentity = BuildEntryIdentity (text2, text3),
				Status = status,
				ReviewedOn = reviewedOn,
				Reviewer = reviewer,
				Notes = notes
			};
		}

		private static void AddEntryIfValid (ICollection<DocumentExecutionPolicyEntry> entries, DocumentExecutionPolicyEntry entry)
		{
			if (entries != null && entry != null) {
				string normalizedKey = (entry.Key ?? string.Empty).Trim ();
				string normalizedTemplateFileName = (entry.TemplateFileName ?? string.Empty).Trim ();
				if (normalizedKey.Length != 0 && normalizedTemplateFileName.Length != 0 && !entries.Any ((DocumentExecutionPolicyEntry item) => item != null && string.Equals (item.Key, normalizedKey, StringComparison.OrdinalIgnoreCase) && string.Equals (item.TemplateFileName, normalizedTemplateFileName, StringComparison.OrdinalIgnoreCase))) {
					entries.Add (new DocumentExecutionPolicyEntry {
						Key = normalizedKey,
						TemplateFileName = normalizedTemplateFileName
					});
				}
			}
		}

		private static string BuildEntryIdentity (string key, string templateFileName)
		{
			string text = (key ?? string.Empty).Trim ();
			string text2 = (templateFileName ?? string.Empty).Trim ();
			return text + "|" + text2;
		}

		private void LogReviewStatusConsistency (IReadOnlyDictionary<string, DocumentReviewStatusSummary> reviewStatusByIdentity, string reviewNotesPath)
		{
			if (reviewStatusByIdentity != null && reviewStatusByIdentity.Count != 0) {
				List<string> list = (from pair in reviewStatusByIdentity
					where pair.Value != null && pair.Value.HasConflictingKnownStatuses
					select FormatReviewStatus (pair.Key, pair.Value)).OrderBy ((string value) => value, StringComparer.OrdinalIgnoreCase).ToList ();
				if (list.Count > 0) {
					_logger.Warn ("Document execution review notes contain conflicting known statuses. keys=" + string.Join (",", list) + ", reviewNotesPath=" + reviewNotesPath);
				}
				List<string> list2 = (from pair in reviewStatusByIdentity
					where pair.Value != null && pair.Value.HasDuplicateStatuses
					select FormatReviewStatus (pair.Key, pair.Value)).OrderBy ((string value) => value, StringComparer.OrdinalIgnoreCase).ToList ();
				if (list2.Count > 0) {
					_logger.Warn ("Document execution review notes contain duplicate statuses. keys=" + string.Join (",", list2) + ", reviewNotesPath=" + reviewNotesPath);
				}
			}
		}

		private static string FormatReviewStatus (string entryIdentity, DocumentReviewStatusSummary summary)
		{
			DocumentReviewStatusSummary documentReviewStatusSummary = summary ?? new DocumentReviewStatusSummary ();
			return (entryIdentity ?? string.Empty) + "(PASS=" + documentReviewStatusSummary.PassCount + ",HOLD=" + documentReviewStatusSummary.HoldCount + ",FAIL=" + documentReviewStatusSummary.FailCount + ",OTHER=" + documentReviewStatusSummary.OtherCount + ")";
		}

		private void ValidateReviewEntryMetadata (ReviewEntryInfo reviewEntry, int lineNumber, string reviewNotesPath)
		{
			if (reviewEntry != null) {
				if (!IsValidReviewDate (reviewEntry.ReviewedOn)) {
					_logger.Warn ("Document execution review notes contain invalid reviewedOn. entry=" + reviewEntry.EntryIdentity + ", line=" + lineNumber + ", value=" + (reviewEntry.ReviewedOn ?? string.Empty) + ", expectedFormat=yyyy-MM-dd, reviewNotesPath=" + reviewNotesPath);
				}
				if (IsPlaceholderMetadataValue (reviewEntry.Reviewer, "reviewer")) {
					_logger.Warn ("Document execution review notes contain placeholder reviewer. entry=" + reviewEntry.EntryIdentity + ", line=" + lineNumber + ", reviewNotesPath=" + reviewNotesPath);
				}
				if (IsPlaceholderMetadataValue (reviewEntry.Notes, "notes")) {
					_logger.Warn ("Document execution review notes contain placeholder notes. entry=" + reviewEntry.EntryIdentity + ", line=" + lineNumber + ", reviewNotesPath=" + reviewNotesPath);
				}
			}
		}

		private static bool IsValidReviewDate (string reviewedOn)
		{
			string text = (reviewedOn ?? string.Empty).Trim ();
			if (text.Length == 0) {
				return false;
			}
			DateTime result;
			return DateTime.TryParseExact (text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
		}

		private static bool IsPlaceholderMetadataValue (string value, string placeholder)
		{
			string text = (value ?? string.Empty).Trim ();
			return text.Length == 0 || string.Equals (text, placeholder ?? string.Empty, StringComparison.OrdinalIgnoreCase);
		}
	}
}
