using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class AccountingNavigationControl : UserControl, ITaskPaneView
	{
		private static readonly Color PaneBackColor = Color.FromArgb (242, 242, 242);

		private static readonly Color DefaultButtonColor = ColorTranslator.FromHtml ("#FFFCF5");

		private static readonly Color HoverButtonColor = ColorTranslator.FromHtml ("#BFEDF8");

		private static readonly Color PressedButtonColor = ColorTranslator.FromHtml ("#FFF1DA");

		private static readonly Color DefaultBorderColor = Color.DeepSkyBlue;

		private static readonly Color HoverBorderColor = Color.DeepSkyBlue;

		private const int PreferredPaneWidth = 292;

		private readonly FlowLayoutPanel _rootPanel;

		private string _lastRenderSignature;

		public int PreferredWidth => 292;

		internal event EventHandler<AccountingNavigationActionEventArgs> ActionInvoked;

		internal AccountingNavigationControl ()
		{
			Dock = DockStyle.Fill;
			BackColor = PaneBackColor;
			_rootPanel = new FlowLayoutPanel {
				Dock = DockStyle.Fill,
				FlowDirection = FlowDirection.TopDown,
				WrapContents = false,
				AutoScroll = true,
				Padding = new Padding (12, 14, 12, 14),
				BackColor = BackColor
			};
			base.Controls.Add (_rootPanel);
		}

		internal void Render (IReadOnlyList<AccountingNavigationActionDefinition> actions)
		{
			string text = BuildSignature (actions);
			if (string.Equals (_lastRenderSignature, text, StringComparison.Ordinal)) {
				return;
			}
			_rootPanel.SuspendLayout ();
			_rootPanel.Controls.Clear ();
			if (actions == null || actions.Count == 0) {
				_rootPanel.Controls.Add (new Label {
					AutoSize = false,
					Width = 216,
					Height = 44,
					Text = "このシートで利用できる会計操作はありません。",
					Font = new Font ("Yu Gothic UI", 9f, FontStyle.Regular),
					ForeColor = Color.FromArgb (90, 90, 90)
				});
			} else {
				foreach (IGrouping<string, AccountingNavigationActionDefinition> item in from a in actions
					group a by a.SectionTitle) {
					_rootPanel.Controls.Add (CreateSectionLabel (item.Key));
					foreach (AccountingNavigationActionDefinition item2 in item) {
						_rootPanel.Controls.Add (CreateActionButton (item2));
					}
				}
			}
			_rootPanel.ResumeLayout ();
			_lastRenderSignature = text;
		}

		private static string BuildSignature (IReadOnlyList<AccountingNavigationActionDefinition> actions)
		{
			if (actions == null || actions.Count == 0) {
				return string.Empty;
			}
			StringBuilder stringBuilder = new StringBuilder ();
			foreach (AccountingNavigationActionDefinition action in actions) {
				stringBuilder.Append (action.SectionTitle ?? string.Empty).Append ('|').Append (action.ActionId ?? string.Empty)
					.Append ('|')
					.Append (action.Caption ?? string.Empty)
					.Append ('|')
					.Append (action.IsEnabled ? '1' : '0')
					.Append (';');
			}
			return stringBuilder.ToString ();
		}

		private Control CreateSectionLabel (string text)
		{
			return new Label {
				AutoSize = false,
				Width = 216,
				Height = 22,
				Margin = new Padding (0, 0, 0, 6),
				Font = new Font ("Yu Gothic UI", 10f, FontStyle.Bold),
				ForeColor = Color.FromArgb (70, 70, 70),
				Text = (text ?? string.Empty)
			};
		}

		private Control CreateActionButton (AccountingNavigationActionDefinition action)
		{
			Button button = new Button {
				Width = 216,
				Height = 36,
				Margin = new Padding (0, 0, 0, 8),
				Text = action.Caption,
				Enabled = action.IsEnabled,
				Tag = action.ActionId,
				FlatStyle = FlatStyle.Flat,
				BackColor = DefaultButtonColor,
				Cursor = Cursors.Hand
			};
			button.FlatAppearance.BorderColor = DefaultBorderColor;
			button.FlatAppearance.MouseOverBackColor = HoverButtonColor;
			button.FlatAppearance.MouseDownBackColor = PressedButtonColor;
			button.Click += ActionButton_Click;
			button.MouseEnter += delegate {
				button.FlatAppearance.BorderColor = HoverBorderColor;
			};
			button.MouseLeave += delegate {
				button.FlatAppearance.BorderColor = DefaultBorderColor;
			};
			return button;
		}

		private void ActionButton_Click (object sender, EventArgs e)
		{
			string text = ((!(sender is Button button)) ? string.Empty : (button.Tag as string));
			EventHandler<AccountingNavigationActionEventArgs> eventHandler = this.ActionInvoked;
			if (eventHandler != null && !string.IsNullOrWhiteSpace (text)) {
				eventHandler (this, new AccountingNavigationActionEventArgs (text));
			}
		}
	}
}
