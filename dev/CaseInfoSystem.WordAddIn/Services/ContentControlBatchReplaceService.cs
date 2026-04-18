using System;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace CaseInfoSystem.WordAddIn.Services
{
	internal sealed class ContentControlBatchReplaceService
	{
		internal sealed class ReplaceRequest
		{
			public string OldTag { get; set; }

			public string NewTag { get; set; }

			public string OldTitle { get; set; }

			public string NewTitle { get; set; }

			public bool UsePartialMatch { get; set; }
		}

		internal sealed class ReplaceResult
		{
			public int ScannedCount { get; set; }

			public int TagChangedCount { get; set; }

			public int TitleChangedCount { get; set; }
		}

		internal sealed class NextReplaceResult
		{
			public bool FoundMatch { get; set; }

			public int TagChangedCount { get; set; }

			public int TitleChangedCount { get; set; }

			public int ControlStart { get; set; }

			public string ControlTag { get; set; }

			public string ControlTitle { get; set; }
		}

		public ReplaceResult Execute (Document document, ReplaceRequest request)
		{
			if (document == null) {
				throw new ArgumentNullException ("document");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			ReplaceResult replaceResult = new ReplaceResult ();
			foreach (ContentControl contentControl in document.ContentControls) {
				if (IsSupportedType (contentControl)) {
					replaceResult.ScannedCount++;
					string text = ReplaceValue (contentControl.Tag, request.OldTag, request.NewTag, request.UsePartialMatch);
					if (!string.Equals (text, contentControl.Tag, StringComparison.Ordinal)) {
						contentControl.Tag = text;
						replaceResult.TagChangedCount++;
					}
					string text2 = ReplaceValue (contentControl.Title, request.OldTitle, request.NewTitle, request.UsePartialMatch);
					if (!string.Equals (text2, contentControl.Title, StringComparison.Ordinal)) {
						contentControl.Title = text2;
						replaceResult.TitleChangedCount++;
					}
				}
			}
			return replaceResult;
		}

		public NextReplaceResult ExecuteNextFromSelection (Document document, Selection selection, ReplaceRequest request)
		{
			if (document == null) {
				throw new ArgumentNullException ("document");
			}
			if (selection == null) {
				throw new ArgumentNullException ("selection");
			}
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			int start = selection.Range.Start;
			ContentControl contentControl = null;
			foreach (ContentControl contentControl2 in document.ContentControls) {
				if (!IsSupportedType (contentControl2) || !IsAtOrAfterSelection (contentControl2, start) || !MatchesTarget (contentControl2, request)) {
					continue;
				}
				contentControl = contentControl2;
				break;
			}
			if (contentControl == null) {
				return new NextReplaceResult {
					FoundMatch = false
				};
			}
			string text = ReplaceValue (contentControl.Tag, request.OldTag, request.NewTag, request.UsePartialMatch);
			string text2 = ReplaceValue (contentControl.Title, request.OldTitle, request.NewTitle, request.UsePartialMatch);
			int tagChangedCount = 0;
			if (!string.Equals (text, contentControl.Tag, StringComparison.Ordinal)) {
				contentControl.Tag = text;
				tagChangedCount = 1;
			}
			int titleChangedCount = 0;
			if (!string.Equals (text2, contentControl.Title, StringComparison.Ordinal)) {
				contentControl.Title = text2;
				titleChangedCount = 1;
			}
			int end = contentControl.Range.End;
			selection.SetRange (end, end);
			return new NextReplaceResult {
				FoundMatch = true,
				TagChangedCount = tagChangedCount,
				TitleChangedCount = titleChangedCount,
				ControlStart = contentControl.Range.Start,
				ControlTag = (contentControl.Tag ?? string.Empty),
				ControlTitle = (contentControl.Title ?? string.Empty)
			};
		}

		public static bool HasAnyTarget (ReplaceRequest request)
		{
			if (request == null) {
				return false;
			}
			return !string.IsNullOrEmpty (request.OldTag) || !string.IsNullOrEmpty (request.OldTitle);
		}

		public static string BuildCompletionMessage (ReplaceResult result)
		{
			if (result == null) {
				return "置換結果を取得できませんでした。";
			}
			StringBuilder stringBuilder = new StringBuilder ();
			stringBuilder.AppendLine ("コンテンツコントロールの置換が完了しました。");
			stringBuilder.AppendLine ("対象コントロール数: " + result.ScannedCount);
			stringBuilder.AppendLine ("Tag 変更数: " + result.TagChangedCount);
			stringBuilder.Append ("Title 変更数: " + result.TitleChangedCount);
			return stringBuilder.ToString ();
		}

		public static string BuildNextReplaceMessage (NextReplaceResult result)
		{
			if (result == null) {
				return "順次置換の結果を取得できませんでした。";
			}
			if (!result.FoundMatch) {
				return "現在の選択位置より下に、一致するコンテンツコントロールはありません。";
			}
			StringBuilder stringBuilder = new StringBuilder ();
			stringBuilder.Append ("1件置換しました");
			stringBuilder.Append (" 位置: " + result.ControlStart);
			stringBuilder.Append (" / Tag変更: " + result.TagChangedCount);
			stringBuilder.Append (" / Title変更: " + result.TitleChangedCount);
			return stringBuilder.ToString ();
		}

		private static bool IsSupportedType (ContentControl contentControl)
		{
			if (contentControl == null) {
				return false;
			}
			return contentControl.Type == WdContentControlType.wdContentControlText || contentControl.Type == WdContentControlType.wdContentControlRichText;
		}

		private static bool IsAtOrAfterSelection (ContentControl contentControl, int selectionStart)
		{
			if (contentControl == null) {
				return false;
			}
			int start = contentControl.Range.Start;
			int end = contentControl.Range.End;
			return start >= selectionStart || (start < selectionStart && selectionStart < end);
		}

		private static bool MatchesTarget (ContentControl contentControl, ReplaceRequest request)
		{
			if (contentControl == null || request == null) {
				return false;
			}
			bool flag = IsMatch (contentControl.Tag, request.OldTag, request.UsePartialMatch);
			bool flag2 = IsMatch (contentControl.Title, request.OldTitle, request.UsePartialMatch);
			return flag || flag2;
		}

		private static bool IsMatch (string source, string oldValue, bool usePartialMatch)
		{
			if (string.IsNullOrEmpty (oldValue)) {
				return false;
			}
			string text = source ?? string.Empty;
			if (usePartialMatch) {
				return text.IndexOf (oldValue, StringComparison.Ordinal) >= 0;
			}
			return string.Equals (text, oldValue, StringComparison.Ordinal);
		}

		private static string ReplaceValue (string source, string oldValue, string newValue, bool usePartialMatch)
		{
			if (string.IsNullOrEmpty (oldValue)) {
				return source ?? string.Empty;
			}
			string text = source ?? string.Empty;
			string text2 = newValue ?? string.Empty;
			if (usePartialMatch) {
				return text.Replace (oldValue, text2);
			}
			return string.Equals (text, oldValue, StringComparison.Ordinal) ? text2 : text;
		}
	}
}
