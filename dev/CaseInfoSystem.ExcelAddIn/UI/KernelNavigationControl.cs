using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal sealed class KernelNavigationControl : UserControl, ITaskPaneView
	{
		private sealed class NavigationButtonSurface : Control
		{
			private bool _isHover;

			private bool _isPressed;

			internal string Caption { get; set; }

			internal bool IsActionEnabled {
				get {
					return Cursor == Cursors.Hand;
				}
				set {
					Cursor = (value ? Cursors.Hand : Cursors.Default);
				}
			}

			internal bool IsCurrentDisplay { get; set; }

			internal NavigationButtonSurface ()
			{
				SetStyle (ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, value: true);
				BackColor = DefaultButtonColor;
				Font = new Font ("Yu Gothic UI", 10f, FontStyle.Regular);
				Cursor = Cursors.Hand;
				base.TabStop = false;
				base.AccessibleRole = AccessibleRole.PushButton;
			}

			protected override void OnPaint (PaintEventArgs e)
			{
				base.OnPaint (e);
				Graphics graphics = e.Graphics;
				Rectangle clientRectangle = base.ClientRectangle;
				Color color = BackColor;
				if (IsCurrentDisplay) {
					color = CurrentDisplayButtonColor;
				} else if (IsActionEnabled && _isPressed) {
					color = PressedButtonColor;
				} else if (IsActionEnabled && _isHover) {
					color = HoverButtonColor;
				}
				Color color2 = (IsCurrentDisplay ? CurrentDisplayBorderColor : ((IsActionEnabled && _isHover) ? HoverBorderColor : DefaultBorderColor));
				Color color3 = (IsCurrentDisplay ? Color.FromArgb (130, 130, 130) : Color.FromArgb (30, 30, 30));
				using (SolidBrush brush = new SolidBrush (color)) {
					using (Pen pen = new Pen (color2, 1f)) {
						using (SolidBrush brush2 = new SolidBrush (color3)) {
							using (StringFormat stringFormat = new StringFormat ()) {
								graphics.FillRectangle (brush, clientRectangle);
								graphics.DrawRectangle (pen, 0, 0, clientRectangle.Width - 1, clientRectangle.Height - 1);
								stringFormat.Alignment = StringAlignment.Center;
								stringFormat.LineAlignment = StringAlignment.Center;
								graphics.DrawString (Caption ?? string.Empty, Font, brush2, clientRectangle, stringFormat);
							}
						}
					}
				}
			}

			protected override void OnMouseEnter (EventArgs e)
			{
				base.OnMouseEnter (e);
				_isHover = true;
				Invalidate ();
			}

			protected override void OnMouseLeave (EventArgs e)
			{
				base.OnMouseLeave (e);
				_isHover = false;
				_isPressed = false;
				Invalidate ();
			}

			protected override void OnMouseDown (MouseEventArgs e)
			{
				base.OnMouseDown (e);
				if (IsActionEnabled && e.Button == MouseButtons.Left) {
					_isPressed = true;
					Invalidate ();
				}
			}

			protected override void OnMouseUp (MouseEventArgs e)
			{
				base.OnMouseUp (e);
				if (_isPressed) {
					_isPressed = false;
					Invalidate ();
				}
			}
		}

		private static readonly Color PaneBackColor = Color.FromArgb (242, 242, 242);

		private static readonly Color DefaultButtonColor = ColorTranslator.FromHtml ("#FFFCF5");

		private static readonly Color HoverButtonColor = ColorTranslator.FromHtml ("#BFEDF8");

		private static readonly Color PressedButtonColor = ColorTranslator.FromHtml ("#FFF1DA");

		private static readonly Color CurrentDisplayButtonColor = Color.FromArgb (245, 245, 245);

		private static readonly Color DefaultBorderColor = Color.DeepSkyBlue;

		private static readonly Color HoverBorderColor = Color.DeepSkyBlue;

		private static readonly Color CurrentDisplayBorderColor = Color.DeepSkyBlue;

		private readonly FlowLayoutPanel _rootPanel;

		private string _lastRenderSignature;

		public int PreferredWidth => 260;

		internal event EventHandler<KernelNavigationActionEventArgs> ActionInvoked;

		internal KernelNavigationControl ()
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

		internal void Render (IReadOnlyList<KernelNavigationActionDefinition> actions)
		{
			string text = BuildSignature (actions);
			if (string.Equals (_lastRenderSignature, text, StringComparison.Ordinal)) {
				return;
			}
			_rootPanel.SuspendLayout ();
			_rootPanel.Controls.Clear ();
			foreach (IGrouping<string, KernelNavigationActionDefinition> item in from a in actions
				group a by a.SectionTitle) {
				_rootPanel.Controls.Add (CreateSectionLabel (item.Key));
				foreach (KernelNavigationActionDefinition item2 in item) {
					_rootPanel.Controls.Add (CreateActionButton (item2));
				}
			}
			_rootPanel.ResumeLayout ();
			_lastRenderSignature = text;
		}

		private static string BuildSignature (IReadOnlyList<KernelNavigationActionDefinition> actions)
		{
			if (actions == null || actions.Count == 0) {
				return string.Empty;
			}
			StringBuilder stringBuilder = new StringBuilder ();
			foreach (KernelNavigationActionDefinition action in actions) {
				stringBuilder.Append (action.SectionTitle ?? string.Empty).Append ('|').Append (action.ActionId ?? string.Empty)
					.Append ('|')
					.Append (action.Caption ?? string.Empty)
					.Append ('|')
					.Append (action.IsEnabled ? '1' : '0')
					.Append ('|')
					.Append (action.IsCurrentDisplay ? '1' : '0')
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

		private Control CreateActionButton (KernelNavigationActionDefinition action)
		{
			NavigationButtonSurface navigationButtonSurface = new NavigationButtonSurface {
				Width = 216,
				Height = 36,
				Margin = new Padding (0, 0, 0, 8),
				Tag = action.ActionId,
				Name = (action.ActionId ?? string.Empty),
				Text = (action.Caption ?? string.Empty),
				Caption = action.Caption,
				IsActionEnabled = action.IsEnabled,
				IsCurrentDisplay = action.IsCurrentDisplay
			};
			navigationButtonSurface.AccessibleName = action.Caption ?? string.Empty;
			navigationButtonSurface.AccessibleRole = AccessibleRole.PushButton;
			navigationButtonSurface.AccessibleDescription = (action.IsEnabled ? "enabled" : "disabled");
			if (action.IsEnabled) {
				navigationButtonSurface.Click += ActionSurface_Click;
			}
			return navigationButtonSurface;
		}

		private void ActionSurface_Click (object sender, EventArgs e)
		{
			if (sender is Control control && !string.Equals (control.AccessibleDescription, "disabled", StringComparison.Ordinal)) {
				string text = control.Tag as string;
				if (string.IsNullOrWhiteSpace (text) && control.Parent != null) {
					text = control.Parent.Tag as string;
				}
				EventHandler<KernelNavigationActionEventArgs> eventHandler = this.ActionInvoked;
				if (eventHandler != null && !string.IsNullOrWhiteSpace (text)) {
					eventHandler (this, new KernelNavigationActionEventArgs (text));
				}
			}
		}
	}
}
