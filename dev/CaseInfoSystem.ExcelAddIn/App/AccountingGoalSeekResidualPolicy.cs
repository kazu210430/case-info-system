using System;
using System.Globalization;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class AccountingGoalSeekResidualPolicy
	{
		internal const double AllowedResidualYenExclusive = 1.0;

		internal static bool IsWithinAllowedResidual (double current, double target)
		{
			double residualAbs = GetResidualAbs (current, target);
			return IsFinite (residualAbs) && residualAbs < AllowedResidualYenExclusive;
		}

		internal static bool ShouldShowResidualNotice (double current, double target)
		{
			return !IsWithinAllowedResidual (current, target);
		}

		internal static double GetResidual (double current, double target)
		{
			return current - target;
		}

		internal static double GetResidualAbs (double current, double target)
		{
			return Math.Abs (GetResidual (current, target));
		}

		internal static string CreateResidualNoticeUserMessage (double current, double target)
		{
			return FormatResidualYen (GetResidualAbs (current, target)) + "円の誤差が生じています。入力内容をご確認ください。";
		}

		internal static string FormatResidualYen (double residualAbs)
		{
			if (!IsFinite (residualAbs)) {
				return "1";
			}

			double yen = Math.Floor (residualAbs);
			if (yen < 1) {
				yen = 1;
			}
			if (yen > long.MaxValue) {
				return yen.ToString ("0", CultureInfo.InvariantCulture);
			}

			return ((long)yen).ToString (CultureInfo.InvariantCulture);
		}

		private static bool IsFinite (double value)
		{
			return !double.IsNaN (value) && !double.IsInfinity (value);
		}
	}
}
