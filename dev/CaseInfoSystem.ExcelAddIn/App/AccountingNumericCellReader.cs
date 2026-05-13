using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingNumericCellReader
	{
		private const string RequiredSupplement = "数値または数式結果の数値を設定してください。空欄は許容されません。";

		private const string AllowBlankAsZeroSupplement = "空欄は 0 として扱えますが、現在の値は空欄ではありません。数値または空欄を設定してください。";

		internal static bool TryParseNumericCell (object cellValue, string displayText, out double value, out bool isBlank)
		{
			value = 0.0;
			isBlank = IsBlankCellValue (cellValue);
			if (isBlank) {
				return false;
			}
			string text = (displayText ?? string.Empty).Trim ();
			if (LooksLikeExcelErrorDisplay (text) || cellValue is ErrorWrapper || cellValue is bool) {
				return false;
			}
			if (cellValue is string text2) {
				text2 = text2.Trim ();
				if (text2.Length == 0) {
					isBlank = true;
					return false;
				}
				return double.TryParse (text2, NumberStyles.Number, CultureInfo.InvariantCulture, out value) || double.TryParse (text2, NumberStyles.Number, CultureInfo.CurrentCulture, out value);
			}
			try {
				value = Convert.ToDouble (cellValue, CultureInfo.InvariantCulture);
				return true;
			} catch {
				return false;
			}
		}

		internal static InvalidOperationException CreateReadFailureException (string sheetName, string cellAddress, string itemName, string procedureName, string displayText, bool allowBlankAsZero)
		{
			string text = string.IsNullOrWhiteSpace (displayText) ? "（空欄）" : displayText.Trim ();
			string text2 = allowBlankAsZero ? AllowBlankAsZeroSupplement : RequiredSupplement;
			string address = cellAddress ?? string.Empty;
			return new InvalidOperationException ((sheetName ?? string.Empty) + "シートのセル " + address + "（" + (itemName ?? string.Empty) + "）の数値読取に失敗しました。処理: " + (procedureName ?? string.Empty) + "。セル番地: " + address + "。セル表示値: " + text + "。" + text2);
		}

		private static bool IsBlankCellValue (object cellValue)
		{
			if (cellValue == null) {
				return true;
			}
			string text = cellValue as string;
			return text != null && string.IsNullOrWhiteSpace (text);
		}

		private static bool LooksLikeExcelErrorDisplay (string displayText)
		{
			return !string.IsNullOrWhiteSpace (displayText) && displayText.StartsWith ("#", StringComparison.Ordinal);
		}
	}
}
