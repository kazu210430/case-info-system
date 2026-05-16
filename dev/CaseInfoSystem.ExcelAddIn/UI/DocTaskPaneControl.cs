using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    internal sealed class DocTaskPaneControl : UserControl
    {
        private sealed class TabButton : Control
        {
            private bool _isHovering;

            internal Color BaseBackColor { get; set; }

            internal bool IsSelected { get; set; }

            internal ContentAlignment TextAlign { get; set; }

            internal new Padding Padding { get; set; }

            public TabButton()
            {
                Cursor = Cursors.Hand;
                ForeColor = Color.FromArgb(40, 40, 40);
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, value: true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle rect = new Rectangle(0, 0, base.Width - 1, base.Height - 1);
                Color fillColor = IsSelected ? GetSelectedFillColor(BaseBackColor) : (_isHovering ? GetHoverFillColor(BaseBackColor) : BaseBackColor);
                Color borderColor = IsSelected ? Color.DeepSkyBlue : (_isHovering ? HoverButtonBorderColor : DefaultButtonBorderColor);
                int borderWidth = 1;
                using (GraphicsPath path = CreateStaticLeftRoundedPath(rect, 12))
                using (SolidBrush brush = new SolidBrush(fillColor))
                using (Pen pen = new Pen(borderColor, borderWidth))
                {
                    e.Graphics.FillPath(brush, path);
                    e.Graphics.DrawPath(pen, path);
                }

                TextFormatFlags flags = TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter;
                flags = TextAlign == ContentAlignment.MiddleCenter
                    ? flags | TextFormatFlags.HorizontalCenter
                    : (TextAlign == ContentAlignment.MiddleRight ? flags | TextFormatFlags.Right : flags | TextFormatFlags.Default);
                Rectangle bounds = new Rectangle(Padding.Left, 0, base.Width - Padding.Left - Padding.Right, base.Height);
                TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor, flags);
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                _isHovering = true;
                Invalidate();
                base.OnMouseEnter(e);
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                _isHovering = false;
                Invalidate();
                base.OnMouseLeave(e);
            }

            private static GraphicsPath CreateStaticLeftRoundedPath(Rectangle rect, int radius)
            {
                GraphicsPath path = new GraphicsPath();
                int diameter = radius * 2;
                path.StartFigure();
                path.AddArc(rect.X, rect.Y, diameter, diameter, 180f, 90f);
                path.AddLine(rect.X + radius, rect.Y, rect.Right, rect.Y);
                path.AddLine(rect.Right, rect.Y, rect.Right, rect.Bottom);
                path.AddLine(rect.Right, rect.Bottom, rect.X + radius, rect.Bottom);
                path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90f, 90f);
                path.CloseFigure();
                return path;
            }

            private static Color GetSelectedFillColor(Color baseColor)
            {
                return baseColor.ToArgb() == DefaultDocumentButtonColor.ToArgb() ? DefaultHoverButtonColor : ControlPaint.Light(baseColor);
            }

            private static Color GetHoverFillColor(Color baseColor)
            {
                return baseColor.ToArgb() == DefaultDocumentButtonColor.ToArgb() ? DefaultHoverButtonColor : GetHoverAccentFillColor(baseColor);
            }
        }

        private sealed class ActionButton : Control
        {
            private bool _isHovering;

            internal Color FillColor { get; set; }

            internal Color BorderColor { get; set; }

            internal ContentAlignment TextAlign { get; set; }

            internal new Padding Padding { get; set; }

            public ActionButton()
            {
                Cursor = Cursors.Hand;
                ForeColor = Color.FromArgb(40, 40, 40);
                DoubleBuffered = true;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, value: true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                Rectangle rect = new Rectangle(0, 0, base.Width - 1, base.Height - 1);
                Color fillColor = _isHovering ? GetHoverFillColor(FillColor) : FillColor;
                Color borderColor = _isHovering ? HoverButtonBorderColor : BorderColor;
                using (SolidBrush brush = new SolidBrush(fillColor))
                using (Pen pen = new Pen(borderColor, 1f))
                {
                    e.Graphics.FillRectangle(brush, rect);
                    e.Graphics.DrawRectangle(pen, rect);
                }

                TextFormatFlags flags = TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter;
                flags = TextAlign == ContentAlignment.MiddleRight
                    ? flags | TextFormatFlags.Right
                    : (TextAlign == ContentAlignment.MiddleCenter ? flags | TextFormatFlags.HorizontalCenter : flags | TextFormatFlags.Default);
                Rectangle bounds = new Rectangle(Padding.Left, 0, base.Width - Padding.Left - Padding.Right, base.Height);
                TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor, flags);
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                _isHovering = true;
                Invalidate();
                base.OnMouseEnter(e);
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                _isHovering = false;
                Invalidate();
                base.OnMouseLeave(e);
            }

            private static Color GetHoverFillColor(Color baseColor)
            {
                return baseColor.ToArgb() == DefaultDocumentButtonColor.ToArgb() ? DefaultHoverButtonColor : GetHoverAccentFillColor(baseColor);
            }
        }

        private sealed class ActionButtonTag
        {
            internal string ActionKind { get; set; }

            internal string Key { get; set; }
        }

        private static readonly Color PaneBackColor = Color.MintCream;
        private static readonly Color DefaultDocumentButtonColor = ColorTranslator.FromHtml("#FFFCF5");
        private static readonly Color DefaultHoverButtonColor = ColorTranslator.FromHtml("#BFEDF8");
        private static readonly Color DefaultButtonBorderColor = Color.DeepSkyBlue;
        private static readonly Color HoverButtonBorderColor = Color.DeepSkyBlue;
        private static readonly Color NoticeBackColor = PaneBackColor;
        private static readonly Color NoticeForeColor = Color.FromArgb(120, 120, 120);

        private readonly Panel _rootPanel;
        private readonly Label _statusLabel;
        private readonly Panel _noticePanel;
        private readonly Label _noticeLabel;
        private CaseTaskPaneViewState _currentViewState;

        private static Color GetHoverAccentFillColor(Color baseColor)
        {
            return BlendColor(baseColor, Color.DeepSkyBlue, 0.25f);
        }

        private static Color BlendColor(Color fromColor, Color toColor, float amount)
        {
            if (amount <= 0f)
            {
                return fromColor;
            }

            if (amount >= 1f)
            {
                return toColor;
            }

            int red = fromColor.R + (int)Math.Round((toColor.R - fromColor.R) * amount);
            int green = fromColor.G + (int)Math.Round((toColor.G - fromColor.G) * amount);
            int blue = fromColor.B + (int)Math.Round((toColor.B - fromColor.B) * amount);
            return Color.FromArgb(red, green, blue);
        }

        internal DocTaskPaneControl()
        {
            Dock = DockStyle.Fill;
            BackColor = PaneBackColor;
            _rootPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = PaneBackColor
            };
            _statusLabel = new Label
            {
                AutoSize = false,
                Left = 12,
                Top = 12,
                Width = 660,
                Height = 44,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Yu Gothic UI", 10f, FontStyle.Regular),
                ForeColor = Color.FromArgb(80, 80, 80)
            };
            _noticePanel = new Panel
            {
                Left = 12,
                Top = 12,
                Width = 360,
                Height = 28,
                BackColor = NoticeBackColor,
                BorderStyle = BorderStyle.None,
                Visible = false
            };
            _noticeLabel = new Label
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 1, 12, 1),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Yu Gothic UI", 9.5f, FontStyle.Regular),
                ForeColor = NoticeForeColor,
                Text = "※ボタン無反応時はEscキーを押してください。"
            };
            _noticePanel.Controls.Add(_noticeLabel);
            _rootPanel.Controls.Add(_statusLabel);
            Controls.Add(_rootPanel);
            ShowMessage("Select a CASE workbook.");
        }

        internal string SelectedTabName
        {
            get { return _currentViewState == null ? string.Empty : _currentViewState.SelectedTabName; }
        }

        internal int PreferredPaneWidthHint { get; private set; }

        internal event EventHandler<TaskPaneActionEventArgs> ActionInvoked;

        internal void ShowMessage(string message)
        {
            _currentViewState = null;
            PreferredPaneWidthHint = 0;
            _rootPanel.SuspendLayout();
            _rootPanel.Controls.Clear();
            _noticePanel.Visible = false;
            _statusLabel.Text = message ?? string.Empty;
            _rootPanel.Controls.Add(_statusLabel);
            _rootPanel.ResumeLayout();
        }

        internal void Render(CaseTaskPaneViewState viewState)
        {
            _currentViewState = viewState;
            _rootPanel.SuspendLayout();
            _rootPanel.Controls.Clear();
            if (viewState == null)
            {
                _rootPanel.ResumeLayout();
                ShowMessage("Failed to load definitions.");
                return;
            }

            if (viewState.HasStatusMessage)
            {
                _rootPanel.ResumeLayout();
                ShowMessage(viewState.StatusMessage);
                return;
            }

            IList<CaseTaskPaneTabPageViewState> orderedTabs = GetOrderedTabs(viewState);
            CaseTaskPaneTabPageViewState selectedTabPage = viewState.GetSelectedTabPage();
            if (orderedTabs.Count == 0 || selectedTabPage == null)
            {
                _rootPanel.ResumeLayout();
                ShowMessage("No available document buttons.");
                return;
            }

            int tabWidth = CalculateTabWidth(orderedTabs);
            int buttonWidth = CalculateButtonWidth(viewState);
            int contentLeft = 12 + tabWidth + 16;
            int startTop = CalculateDocumentTop(viewState.SpecialButtons);
            // 処理ブロック: ここから先はフォント幅やスクロール幅に依存する純表示レイアウトだけを扱う。
            PreferredPaneWidthHint = CalculatePreferredPaneWidth(tabWidth, buttonWidth);
            AddNoticeBand();
            AddTabButtons(orderedTabs, tabWidth, startTop);
            AddSpecialButtons(viewState.SpecialButtons, contentLeft, buttonWidth);
            AddDocumentButtons(selectedTabPage.DocumentButtons, contentLeft, buttonWidth, startTop);
            _rootPanel.ResumeLayout();
        }

        private static IList<CaseTaskPaneTabPageViewState> GetOrderedTabs(CaseTaskPaneViewState viewState)
        {
            return viewState.TabPages.OrderBy(tab => tab.Order).ToList();
        }

        private int CalculateTabWidth(IList<CaseTaskPaneTabPageViewState> tabs)
        {
            int width = 110;
            using (Graphics dc = CreateGraphics())
            using (Font font = new Font("Yu Gothic UI", 10f, FontStyle.Bold))
            {
                foreach (CaseTaskPaneTabPageViewState tab in tabs)
                {
                    width = Math.Max(width, TextRenderer.MeasureText(dc, tab.TabName ?? string.Empty, font).Width + 28);
                }
            }

            return Math.Min(240, width);
        }

        private int CalculateButtonWidth(CaseTaskPaneViewState viewState)
        {
            int width = 160;
            using (Graphics dc = CreateGraphics())
            using (Font font = new Font("Yu Gothic UI", 11f, FontStyle.Regular))
            {
                foreach (CaseTaskPaneTabPageViewState tabPage in viewState.TabPages)
                {
                    foreach (CaseTaskPaneActionViewState documentButton in tabPage.DocumentButtons)
                    {
                        width = Math.Max(width, TextRenderer.MeasureText(dc, documentButton.Caption ?? string.Empty, font).Width + 40);
                    }
                }

                foreach (CaseTaskPaneActionViewState specialButton in viewState.SpecialButtons)
                {
                    width = Math.Max(width, TextRenderer.MeasureText(dc, specialButton.Caption ?? string.Empty, font).Width + 40);
                }
            }

            return Math.Min(420, width);
        }

        private static int CalculateDocumentTop(IReadOnlyList<CaseTaskPaneActionViewState> specialButtons)
        {
            return specialButtons == null || specialButtons.Count == 0 ? 98 : 50 + (48 * specialButtons.Count);
        }

        private static int CalculatePreferredPaneWidth(int tabWidth, int buttonWidth)
        {
            return 12 + tabWidth + 16 + buttonWidth + 12 + SystemInformation.VerticalScrollBarWidth + 8;
        }

        private void AddNoticeBand()
        {
            int textWidth = TextRenderer.MeasureText(_noticeLabel.Text ?? string.Empty, _noticeLabel.Font).Width;
            int desiredWidth = textWidth + _noticeLabel.Padding.Left + _noticeLabel.Padding.Right + 58;
            int paneWidth = Math.Max(PreferredPaneWidthHint, Math.Max(_rootPanel.ClientSize.Width, ClientSize.Width));
            int maxWidth = Math.Max(260, paneWidth - 24 - SystemInformation.VerticalScrollBarWidth);
            _noticePanel.Width = Math.Min(maxWidth, Math.Max(360, desiredWidth));
            int centeredLeft = (paneWidth - _noticePanel.Width - SystemInformation.VerticalScrollBarWidth) / 2;
            _noticePanel.Left = Math.Max(12, centeredLeft);
            _noticePanel.Visible = true;
            _rootPanel.Controls.Add(_noticePanel);
            _noticePanel.BringToFront();
            _noticeLabel.TextAlign = ContentAlignment.MiddleLeft;
        }

        private void AddTabButtons(IList<CaseTaskPaneTabPageViewState> tabs, int tabWidth, int startTop)
        {
            int top = startTop;
            string selectedTabName = SelectedTabName;
            foreach (CaseTaskPaneTabPageViewState tab in tabs)
            {
                bool isSelected = string.Equals(tab.TabName, selectedTabName, StringComparison.Ordinal);
                TabButton tabButton = new TabButton
                {
                    Text = tab.TabName ?? string.Empty,
                    Left = 12,
                    Top = top,
                    Width = tabWidth,
                    Height = 34,
                    Font = new Font("Yu Gothic UI", isSelected ? 10.5f : 10f, isSelected ? FontStyle.Bold : FontStyle.Regular),
                    TextAlign = ContentAlignment.MiddleRight,
                    Padding = new Padding(10, 0, 10, 0),
                    Tag = tab.TabName ?? string.Empty,
                    BaseBackColor = ResolvePaneSurfaceColor(tab.BackColor),
                    IsSelected = isSelected
                };
                tabButton.Click += TabButton_Click;
                _rootPanel.Controls.Add(tabButton);
                top += 42;
            }
        }

        private void AddSpecialButtons(IReadOnlyList<CaseTaskPaneActionViewState> specialButtons, int contentLeft, int buttonWidth)
        {
            if (specialButtons == null || specialButtons.Count == 0)
            {
                return;
            }

            int top = 50;
            foreach (CaseTaskPaneActionViewState specialButton in specialButtons)
            {
                ActionButton actionButton = CreateActionButton(specialButton.Caption, specialButton.BackColor, specialButton.ActionKind, specialButton.Key, buttonWidth, 32);
                actionButton.Left = contentLeft;
                actionButton.Top = top;
                _rootPanel.Controls.Add(actionButton);
                top += 48;
            }
        }

        private void AddDocumentButtons(IReadOnlyList<CaseTaskPaneActionViewState> documentButtons, int contentLeft, int buttonWidth, int startTop)
        {
            int top = startTop;
            foreach (CaseTaskPaneActionViewState documentButton in documentButtons)
            {
                ActionButton actionButton = CreateActionButton(documentButton.Caption, documentButton.BackColor, documentButton.ActionKind, documentButton.Key, buttonWidth, 32);
                actionButton.Left = contentLeft;
                actionButton.Top = top;
                _rootPanel.Controls.Add(actionButton);
                top += 48;
            }

            _rootPanel.AutoScrollMinSize = new Size(0, Math.Max(top, startTop) + 12);
        }

        private ActionButton CreateActionButton(string caption, Color backColor, string actionKind, string key, int width, int height)
        {
            ActionButton actionButton = new ActionButton
            {
                Text = caption ?? string.Empty,
                Width = width,
                Height = height,
                Font = new Font("Yu Gothic UI", 11f, FontStyle.Regular),
                FillColor = ResolvePaneSurfaceColor(backColor),
                BorderColor = DefaultButtonBorderColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(14, 0, 10, 0),
                Tag = new ActionButtonTag
                {
                    ActionKind = actionKind ?? string.Empty,
                    Key = key ?? string.Empty
                }
            };
            actionButton.Click += ActionButton_Click;
            return actionButton;
        }

        private static Color ResolvePaneSurfaceColor(Color backColor)
        {
            return backColor == Color.Empty || backColor.ToArgb() == Color.White.ToArgb() ? DefaultDocumentButtonColor : backColor;
        }

        private void TabButton_Click(object sender, EventArgs e)
        {
            string tabName = (sender as Control)?.Tag as string;
            if (!string.IsNullOrEmpty(tabName) && _currentViewState != null)
            {
                Render(_currentViewState.WithSelectedTab(tabName));
            }
        }

        private void ActionButton_Click(object sender, EventArgs e)
        {
            ActionButtonTag tag = (sender as Control)?.Tag as ActionButtonTag;
            if (tag != null)
            {
                ActionInvoked?.Invoke(this, new TaskPaneActionEventArgs(tag.ActionKind, tag.Key));
            }
        }
    }
}
