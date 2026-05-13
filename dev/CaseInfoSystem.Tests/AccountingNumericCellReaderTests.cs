using System;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class AccountingNumericCellReaderTests
	{
		[Theory]
		[InlineData ("ABC")]
		[InlineData ("--")]
		[InlineData ("1,000円")]
		public void TryParseNumericCell_WhenNonNumericString_ReturnsFalse (string input)
		{
			bool result = AccountingNumericCellReader.TryParseNumericCell (input, input, out double value, out bool isBlank);

			Assert.False (result);
			Assert.False (isBlank);
			Assert.Equal (0.0, value);
		}

		[Theory]
		[InlineData (0.0, "0")]
		[InlineData ("0", "0")]
		[InlineData ("0.0", "0.0")]
		public void TryParseNumericCell_WhenZeroValue_ReturnsTrue (object input, string displayText)
		{
			bool result = AccountingNumericCellReader.TryParseNumericCell (input, displayText, out double value, out bool isBlank);

			Assert.True (result);
			Assert.False (isBlank);
			Assert.Equal (0.0, value);
		}

		[Fact]
		public void TryParseNumericCell_WhenCellValueIsNull_ReturnsBlank ()
		{
			bool result = AccountingNumericCellReader.TryParseNumericCell (null, string.Empty, out double value, out bool isBlank);

			Assert.False (result);
			Assert.True (isBlank);
			Assert.Equal (0.0, value);
		}

		[Theory]
		[InlineData ("")]
		[InlineData ("   ")]
		public void TryParseNumericCell_WhenBlankString_ReturnsBlank (string input)
		{
			bool result = AccountingNumericCellReader.TryParseNumericCell (input, input, out double value, out bool isBlank);

			Assert.False (result);
			Assert.True (isBlank);
			Assert.Equal (0.0, value);
		}

		[Fact]
		public void TryParseNumericCell_WhenExcelErrorDisplay_ReturnsFalse ()
		{
			bool result = AccountingNumericCellReader.TryParseNumericCell (2007, "#DIV/0!", out double value, out bool isBlank);

			Assert.False (result);
			Assert.False (isBlank);
			Assert.Equal (0.0, value);
		}

		[Fact]
		public void TryParseNumericCell_WhenFormulaBlankValue_ReturnsBlank ()
		{
			bool result = AccountingNumericCellReader.TryParseNumericCell (string.Empty, string.Empty, out double value, out bool isBlank);

			Assert.False (result);
			Assert.True (isBlank);
			Assert.Equal (0.0, value);
		}

		[Fact]
		public void CreateReadFailureException_WhenRequired_IncludesRequiredSupplement ()
		{
			InvalidOperationException exception = AccountingNumericCellReader.CreateReadFailureException ("請求書", "F23", "請求額小計", "AccountingInstallmentSchedule.LoadFormState", "ABC", allowBlankAsZero: false);

			Assert.Contains ("請求書シートのセル F23（請求額小計）の数値読取に失敗しました。", exception.Message);
			Assert.Contains ("処理: AccountingInstallmentSchedule.LoadFormState。", exception.Message);
			Assert.Contains ("セル番地: F23。", exception.Message);
			Assert.Contains ("セル表示値: ABC。", exception.Message);
			Assert.Contains ("空欄は許容されません。", exception.Message);
		}

		[Fact]
		public void CreateReadFailureException_WhenAllowBlankAsZero_IncludesAllowBlankSupplement ()
		{
			InvalidOperationException exception = AccountingNumericCellReader.CreateReadFailureException ("請求書", "F29", "お預かり金額", "AccountingPaymentHistory.LoadFormState", "--", allowBlankAsZero: true);

			Assert.Contains ("請求書シートのセル F29（お預かり金額）の数値読取に失敗しました。", exception.Message);
			Assert.Contains ("処理: AccountingPaymentHistory.LoadFormState。", exception.Message);
			Assert.Contains ("セル番地: F29。", exception.Message);
			Assert.Contains ("セル表示値: --。", exception.Message);
			Assert.Contains ("数値または空欄を設定してください。", exception.Message);
		}

		[Fact]
		public void CreateReadFailureException_WhenDisplayTextLooksLikeColumnHeader_KeepsCellAddressSeparate ()
		{
			InvalidOperationException exception = AccountingNumericCellReader.CreateReadFailureException ("分割払い予定表", "J13", "実費残高", "AccountingInstallmentSchedule.WriteScheduleRow", "列10", allowBlankAsZero: false);

			Assert.Contains ("分割払い予定表シートのセル J13（実費残高）の数値読取に失敗しました。", exception.Message);
			Assert.Contains ("セル番地: J13。", exception.Message);
			Assert.Contains ("セル表示値: 列10。", exception.Message);
		}
	}
}
