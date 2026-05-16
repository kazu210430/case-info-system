using System.Drawing;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal static class AccountingFormButtonAppearanceHelper
	{
		private static readonly Color TaskPaneDefaultButtonColor = Color.MintCream;

		private static readonly Color TaskPaneHoverButtonColor = ColorTranslator.FromHtml ("#BFEDF8");

		private static readonly Color ButtonBorderColor = Color.DeepSkyBlue;

		internal static void Apply (params Button[] buttons)
		{
			if (buttons == null) {
				return;
			}

			foreach (Button button in buttons) {
				if (button == null) {
					continue;
				}

				ApplySingle (button);
			}
		}

		private static void ApplySingle (Button button)
		{
			Color baseColor = ResolveBaseColor (button);
			if (IsBackgroundLikeButton (button, baseColor)) {
				button.BackColor = TaskPaneDefaultButtonColor;
			}

			button.UseVisualStyleBackColor = false;
			button.FlatStyle = FlatStyle.Flat;
			button.FlatAppearance.BorderColor = ButtonBorderColor;
			button.FlatAppearance.MouseOverBackColor = GetHoverFillColor (baseColor);
			button.FlatAppearance.MouseDownBackColor = GetHoverFillColor (baseColor);
		}

		private static Color ResolveBaseColor (Button button)
		{
			if (button.UseVisualStyleBackColor || button.BackColor.IsEmpty) {
				return SystemColors.Control;
			}

			return button.BackColor;
		}

		private static Color GetHoverFillColor (Color baseColor)
		{
			if (IsDefaultButtonColor (baseColor)) {
				return TaskPaneHoverButtonColor;
			}

			return BlendColor (baseColor, Color.DeepSkyBlue, 0.25f);
		}

		private static bool IsDefaultButtonColor (Color baseColor)
		{
			return baseColor.ToArgb () == SystemColors.Control.ToArgb ()
				|| baseColor.ToArgb () == TaskPaneDefaultButtonColor.ToArgb ();
		}

		private static bool IsBackgroundLikeButton (Button button, Color baseColor)
		{
			return button.UseVisualStyleBackColor
				|| IsNearWhite (baseColor)
				|| IsGrayLike (baseColor);
		}

		private static bool IsNearWhite (Color color)
		{
			return color.R >= 240 && color.G >= 240 && color.B >= 240;
		}

		private static bool IsGrayLike (Color color)
		{
			return System.Math.Abs (color.R - color.G) <= 8
				&& System.Math.Abs (color.G - color.B) <= 8
				&& System.Math.Abs (color.R - color.B) <= 8;
		}

		private static Color BlendColor (Color fromColor, Color toColor, float amount)
		{
			if (amount <= 0f) {
				return fromColor;
			}

			if (amount >= 1f) {
				return toColor;
			}

			int red = fromColor.R + (int)System.Math.Round ((toColor.R - fromColor.R) * amount);
			int green = fromColor.G + (int)System.Math.Round ((toColor.G - fromColor.G) * amount);
			int blue = fromColor.B + (int)System.Math.Round ((toColor.B - fromColor.B) * amount);
			return Color.FromArgb (red, green, blue);
		}
	}
}
