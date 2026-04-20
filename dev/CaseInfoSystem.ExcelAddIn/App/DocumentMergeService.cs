using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class DocumentMergeService
	{
		private sealed class ContentControlEntry
		{
			internal object Control { get; set; }

			internal string Tag { get; set; }

			internal string Title { get; set; }

			internal int Type { get; set; }
		}

		private sealed class ContentControlIndex
		{
			internal IDictionary<string, List<ContentControlEntry>> ByTag { get; private set; }

			internal IDictionary<string, List<ContentControlEntry>> ByTitle { get; private set; }

			internal IDictionary<string, List<ContentControlEntry>> PrimaryKeys { get; private set; }

			internal IDictionary<string, string> PrimarySources { get; private set; }

			internal IDictionary<string, string> DuplicatePrimary { get; private set; }

			internal int EmptyKeyCount { get; set; }

			internal ContentControlIndex ()
			{
				ByTag = new Dictionary<string, List<ContentControlEntry>> (StringComparer.OrdinalIgnoreCase);
				ByTitle = new Dictionary<string, List<ContentControlEntry>> (StringComparer.OrdinalIgnoreCase);
				PrimaryKeys = new Dictionary<string, List<ContentControlEntry>> (StringComparer.OrdinalIgnoreCase);
				PrimarySources = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
				DuplicatePrimary = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			}
		}

		private const int WordContentControlRichText = 0;

		private const int WordContentControlText = 1;

		private const int WordContentControlDate = 6;

		private const int WordContentControlCheckBox = 8;

		private const char WordManualLineBreak = '\v';

		private const string TodayDatePickerTag = "Date";

		private readonly Logger _logger;

		internal DocumentMergeService (Logger logger)
		{
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void ApplyMergeData (object wordDocument, IReadOnlyDictionary<string, string> mergeData)
		{
			if (wordDocument == null) {
				throw new ArgumentNullException ("wordDocument");
			}
			Stopwatch stopwatch = Stopwatch.StartNew ();
			Stopwatch stopwatch2 = Stopwatch.StartNew ();
			ContentControlIndex contentControlIndex = BuildContentControlIndex (wordDocument);
			_logger.Debug ("DocumentMergeService.ApplyMergeData", "IndexBuilt elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " primaryKeyCount=" + contentControlIndex.PrimaryKeys.Count + " emptyKeyCount=" + contentControlIndex.EmptyKeyCount);
			stopwatch2.Restart ();
			List<string> list = new List<string> ();
			List<string> list2 = new List<string> ();
			List<string> list3 = new List<string> ();
			List<string> list4 = new List<string> ();
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int skippedNonTextCount = 0;
			int num5 = 0;
			if (mergeData != null) {
				foreach (KeyValuePair<string, string> mergeDatum in mergeData) {
					num2++;
					List<ContentControlEntry> list5 = FindContentControls (contentControlIndex, mergeDatum.Key);
					if (list5 == null || list5.Count == 0) {
						list.Add (mergeDatum.Key ?? string.Empty);
						continue;
					}
					num3 += list5.Count;
					if (list5.Count > 1) {
						list4.Add ((mergeDatum.Key ?? string.Empty) + " (" + list5.Count + "件)");
					}
					string text = mergeDatum.Value ?? string.Empty;
					string normalizedValueText = NormalizeTextForContentControl (text);
					if (text.Length == 0) {
						num += list5.Count;
					}
					foreach (ContentControlEntry item in list5) {
						if (IsTextContentControl (item)) {
							WritePlainTextToControl (item, normalizedValueText);
							num4++;
						} else {
							list3.Add (BuildControlLabel (mergeDatum.Key ?? string.Empty, item));
							skippedNonTextCount++;
						}
					}
				}
			}
			num5 = ApplySystemDateContentControls (contentControlIndex, list3, ref skippedNonTextCount);
			_logger.Debug ("DocumentMergeService.ApplyMergeData", "ValuesApplied elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " mergeKeyCount=" + (mergeData?.Count ?? 0) + " processedKeyCount=" + num2 + " matchedControlCount=" + num3 + " writtenControlCount=" + num4 + " systemDateControlCount=" + num5 + " skippedNonTextCount=" + skippedNonTextCount);
			stopwatch2.Restart ();
			foreach (KeyValuePair<string, List<ContentControlEntry>> primaryKey in contentControlIndex.PrimaryKeys) {
				if ((mergeData == null || !mergeData.ContainsKey (primaryKey.Key)) && !IsSystemManagedPrimaryKey (primaryKey.Key)) {
					string text2 = (contentControlIndex.PrimarySources.ContainsKey (primaryKey.Key) ? contentControlIndex.PrimarySources [primaryKey.Key] : string.Empty);
					list2.Add (text2 + "=" + primaryKey.Key);
				}
			}
			_logger.Debug ("DocumentMergeService.ApplyMergeData", "MissingCheckCompleted elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " missingInDocumentCount=" + list.Count + " missingInDataCount=" + list2.Count + " duplicateControlsCount=" + list4.Count + " duplicatePrimaryCount=" + contentControlIndex.DuplicatePrimary.Count + " emptyValueCount=" + num);
			stopwatch2.Restart ();
			LogMergeWarnings (list, list2, list3, list4, contentControlIndex.DuplicatePrimary, contentControlIndex.EmptyKeyCount, num);
			_logger.Debug ("DocumentMergeService.ApplyMergeData", "WarningsLogged elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed));
		}

		internal void RemoveContentControlsKeepText (object wordDocument)
		{
			if (wordDocument == null) {
				throw new ArgumentNullException ("wordDocument");
			}
			Stopwatch stopwatch = Stopwatch.StartNew ();
			dynamic val = ((dynamic)wordDocument).ContentControls;
			int num = Convert.ToInt32 (val.Count);
			int num2 = 0;
			int num3 = 0;
			for (int num4 = num; num4 >= 1; num4--) {
				dynamic val2 = val.Item (num4);
				if (ShouldRemoveContentControl (val2)) {
					val2.Delete ();
					num2++;
				} else {
					num3++;
				}
			}
			_logger.Debug ("DocumentMergeService.RemoveContentControlsKeepText", "Completed elapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " initialCount=" + num + " deletedCount=" + num2 + " skippedCheckBoxCount=" + num3);
		}

		private static ContentControlIndex BuildContentControlIndex (object wordDocument)
		{
			dynamic val = ((dynamic)wordDocument).ContentControls;
			ContentControlIndex contentControlIndex = new ContentControlIndex ();
			int num = Convert.ToInt32 (val.Count);
			for (int i = 1; i <= num; i++) {
				dynamic val2 = val.Item (i);
				ContentControlEntry contentControlEntry = new ContentControlEntry
				{
					Control = val2,
					Tag = (Convert.ToString (val2.Tag) ?? string.Empty).Trim (),
					Title = (Convert.ToString (val2.Title) ?? string.Empty).Trim (),
					Type = Convert.ToInt32 (val2.Type)
				};
				string text = contentControlEntry.Tag;
				string text2 = contentControlEntry.Title;
				if (text.Length > 0) {
					DocumentMergeService.AddControlToIndex (contentControlIndex.ByTag, text, contentControlEntry);
				}
				if (text2.Length > 0) {
					DocumentMergeService.AddControlToIndex (contentControlIndex.ByTitle, text2, contentControlEntry);
				}
				string text3;
				string text4;
				if (text.Length > 0) {
					text3 = text;
					text4 = "Tag";
				} else {
					if (text2.Length <= 0) {
						contentControlIndex.EmptyKeyCount++;
						continue;
					}
					text3 = text2;
					text4 = "Title";
				}
				DocumentMergeService.AddControlToIndex (contentControlIndex.PrimaryKeys, text3, contentControlEntry);
				contentControlIndex.PrimarySources [text3] = text4;
				if (contentControlIndex.PrimaryKeys [text3].Count > 1) {
					contentControlIndex.DuplicatePrimary [text3] = text4 + ":" + contentControlIndex.PrimaryKeys [text3].Count;
				}
			}
			return contentControlIndex;
		}

		private static List<ContentControlEntry> FindContentControls (ContentControlIndex index, string keyText)
		{
			if (index == null) {
				return null;
			}
			string text = (keyText ?? string.Empty).Trim ();
			if (text.Length == 0) {
				return null;
			}
			if (index.ByTag.TryGetValue (text, out var value)) {
				return value;
			}
			index.ByTitle.TryGetValue (text, out var value2);
			return value2;
		}

		private static bool IsTextContentControl (ContentControlEntry control)
		{
			int num = control == null ? -1 : control.Type;
			return num == 0 || num == 1;
		}

		private static bool ShouldRemoveContentControl (object control)
		{
			int num = Convert.ToInt32 (((dynamic)control).Type);
			return num != 8 && num != 6;
		}

		private static bool IsDateContentControl (ContentControlEntry control)
		{
			int num = control == null ? -1 : control.Type;
			return num == 6;
		}

		private static int ApplySystemDateContentControls (ContentControlIndex index, IList<string> skippedNonText, ref int skippedNonTextCount)
		{
			if (index == null || !index.ByTag.TryGetValue ("Date", out var value) || value == null || value.Count == 0) {
				return 0;
			}
			int num = 0;
			string normalizedValueText = DateTime.Today.ToString ("yyyy/MM/dd", CultureInfo.InvariantCulture);
			for (int i = 0; i < value.Count; i++) {
				ContentControlEntry control = value [i];
				if (IsDateContentControl (control)) {
					WritePlainTextToControl (control, normalizedValueText);
					num++;
				} else {
					skippedNonText?.Add (BuildControlLabel ("Date", control));
					skippedNonTextCount++;
				}
			}
			return num;
		}

		private static bool IsSystemManagedPrimaryKey (string keyText)
		{
			return string.Equals ((keyText ?? string.Empty).Trim (), "Date", StringComparison.OrdinalIgnoreCase);
		}

		private static void WritePlainTextToControl (ContentControlEntry control, string normalizedValueText)
		{
			if (control == null || control.Control == null) {
				return;
			}
			((dynamic)control.Control).Range.Text = normalizedValueText ?? string.Empty;
		}

		private static string NormalizeTextForContentControl (string valueText)
		{
			if (string.IsNullOrEmpty (valueText)) {
				return string.Empty;
			}
			return valueText.Replace ("\r\n", '\v'.ToString ()).Replace ("\r", '\v'.ToString ()).Replace ("\n", '\v'.ToString ());
		}

		private static string BuildControlLabel (string keyText, ContentControlEntry control)
		{
			return "Key=[" + (keyText ?? string.Empty) + "] Tag=[" + (control == null ? string.Empty : (control.Tag ?? string.Empty)) + "] Title=[" + (control == null ? string.Empty : (control.Title ?? string.Empty)) + "] Type=[" + (control == null ? string.Empty : control.Type.ToString ()) + "]";
		}

		private void LogMergeWarnings (IList<string> missingInDocument, IList<string> missingInData, IList<string> skippedNonText, IList<string> duplicateControls, IDictionary<string, string> duplicatePrimary, int emptyKeyCount, int emptyValueCount)
		{
			StringBuilder stringBuilder = new StringBuilder ();
			AppendBulletSection (stringBuilder, "[文書側に該当CCが無いキー]", missingInDocument);
			AppendBulletSection (stringBuilder, "[差込データに値が無いCC]", missingInData);
			AppendBulletSection (stringBuilder, "[非テキスト系のため未反映]", skippedNonText);
			AppendBulletSection (stringBuilder, "[同一キーへ複数CCを反映]", duplicateControls);
			if (duplicatePrimary != null && duplicatePrimary.Count > 0) {
				stringBuilder.AppendLine ("[文書内で主キー重複]");
				foreach (KeyValuePair<string, string> item in duplicatePrimary) {
					stringBuilder.AppendLine ("・" + item.Key + " (" + item.Value + ")");
				}
			}
			if (emptyKeyCount > 0) {
				stringBuilder.AppendLine ("[Tag/Title 未設定のCC]");
				stringBuilder.AppendLine ("件数: " + emptyKeyCount);
			}
			if (emptyValueCount > 0) {
				stringBuilder.AppendLine ("[空文字列を反映したCC]");
				stringBuilder.AppendLine ("件数: " + emptyValueCount);
			}
			if (stringBuilder.Length == 0) {
				_logger.Info ("DocumentMergeService completed without merge warnings.");
			} else {
				_logger.Info ("DocumentMergeService warnings:" + Environment.NewLine + stringBuilder);
			}
		}

		private static void AppendBulletSection (StringBuilder builder, string title, IList<string> items)
		{
			if (builder != null && items != null && items.Count != 0) {
				builder.AppendLine (title ?? string.Empty);
				for (int i = 0; i < items.Count; i++) {
					builder.AppendLine ("・" + (items [i] ?? string.Empty));
				}
			}
		}

		private static void AddControlToIndex (IDictionary<string, List<ContentControlEntry>> bucket, string keyText, ContentControlEntry control)
		{
			if (!bucket.TryGetValue (keyText, out var value)) {
				value = new List<ContentControlEntry> ();
				bucket.Add (keyText, value);
			}
			value.Add (control);
		}

		private static string FormatElapsedSeconds (TimeSpan elapsed)
		{
			return elapsed.TotalSeconds.ToString ("0.000");
		}
	}
}
