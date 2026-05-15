using System;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class AccountingGoalSeekResidualPolicyTests
	{
		[Theory]
		[InlineData (-0.002117)]
		[InlineData (0.49)]
		[InlineData (-0.99)]
		[InlineData (0.999999)]
		public void IsWithinAllowedResidual_AcceptsResidualsBelowOneYen (double current)
		{
			Assert.True (AccountingGoalSeekResidualPolicy.IsWithinAllowedResidual (current, 0));
			Assert.False (AccountingGoalSeekResidualPolicy.ShouldShowResidualNotice (current, 0));
		}

		[Theory]
		[InlineData (1.0)]
		[InlineData (-1.0)]
		[InlineData (2.3)]
		[InlineData (-2.3)]
		public void ShouldShowResidualNotice_ReturnsTrueForResidualsAtLeastOneYen (double current)
		{
			Assert.False (AccountingGoalSeekResidualPolicy.IsWithinAllowedResidual (current, 0));
			Assert.True (AccountingGoalSeekResidualPolicy.ShouldShowResidualNotice (current, 0));
		}

		[Fact]
		public void CreateResidualNoticeUserMessage_UsesBusinessMessageAndWholeYenDisplay ()
		{
			string message = AccountingGoalSeekResidualPolicy.CreateResidualNoticeUserMessage (2.3, 0);

			Assert.Equal ("2円の誤差が生じています。入力内容をご確認ください。", message);
			Assert.DoesNotContain ("2.3", message);
		}

		[Fact]
		public void CreateResidualNoticeUserMessage_DoesNotExposeTechnicalTerms ()
		{
			string message = AccountingGoalSeekResidualPolicy.CreateResidualNoticeUserMessage (1.0, 0);
			string[] forbiddenTerms = {
				"GoalSeek",
				"formulaCell",
				"changingCell",
				"current",
				"target",
				"residual",
				"InvalidOperationException"
			};

			Assert.Contains ("誤差が生じています", message);
			Assert.Contains ("入力内容をご確認ください", message);
			foreach (string forbiddenTerm in forbiddenTerms) {
				Assert.Equal (-1, message.IndexOf (forbiddenTerm, StringComparison.Ordinal));
			}
		}
	}
}
