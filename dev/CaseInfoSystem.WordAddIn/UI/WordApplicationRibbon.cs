using Microsoft.Office.Tools.Ribbon;

namespace CaseInfoSystem.WordAddIn.UI
{
	internal sealed class WordApplicationRibbon : RibbonBase
	{
		private RibbonTab _tab;

		private RibbonGroup _group;

		private RibbonToggleButton _stylePaneToggleButton;

		private RibbonButton _openBatchReplaceButton;

		public WordApplicationRibbon ()
			: base (Globals.Factory.GetRibbonFactory ())
		{
			InitializeComponent ();
		}

		private void InitializeComponent ()
		{
			_tab = base.Factory.CreateRibbonTab ();
			_group = base.Factory.CreateRibbonGroup ();
			_stylePaneToggleButton = base.Factory.CreateRibbonToggleButton ();
			_openBatchReplaceButton = base.Factory.CreateRibbonButton ();
			_tab.Label = "案件情報System";
			_tab.Groups.Add (_group);
			_group.Label = "Word連携";
			_group.Items.Add (_stylePaneToggleButton);
			_group.Items.Add (_openBatchReplaceButton);
			_stylePaneToggleButton.Label = "Style右表示";
			_stylePaneToggleButton.ScreenTip = "文書ごとに右側のスタイル作業ウィンドウ表示を切り替えます。";
			_stylePaneToggleButton.SuperTip = "この設定は文書内に保存されるため、雛形をコピーして作った文書にも引き継がれます。";
			_stylePaneToggleButton.Click += StylePaneToggleButton_Click;
			_openBatchReplaceButton.Label = "CCタイトル/タグ一括置換";
			_openBatchReplaceButton.ScreenTip = "アクティブ文書のコンテンツコントロールのタイトルとTagを一括置換します。";
			_openBatchReplaceButton.Click += OpenBatchReplaceButton_Click;
			base.Name = "WordApplicationRibbon";
			base.RibbonType = "Microsoft.Word.Document";
			base.Tabs.Add (_tab);
			base.Load += WordApplicationRibbon_Load;
		}

		private void WordApplicationRibbon_Load (object sender, RibbonUIEventArgs e)
		{
			Globals.ThisAddIn?.RegisterRibbon (this);
		}

		private void StylePaneToggleButton_Click (object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn?.ToggleStylePaneForActiveDocument ();
		}

		private void OpenBatchReplaceButton_Click (object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn?.ShowContentControlBatchReplaceForm ();
		}

		internal void SyncStylePaneToggleState (bool enabled, bool hasDocument)
		{
			_stylePaneToggleButton.Enabled = hasDocument;
			_stylePaneToggleButton.Checked = hasDocument && enabled;
		}
	}
}
