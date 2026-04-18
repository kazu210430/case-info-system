using Microsoft.Office.Tools.Ribbon;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    /// <summary>
    /// Class: Excel ribbon entry point for maintenance commands.
    /// Responsibility: render ribbon controls and delegate clicks to ThisAddIn.
    /// </summary>
    internal sealed class CaseInfoSystemRibbon : RibbonBase
    {
        private RibbonTab _tab;
        private RibbonGroup _group;
        private RibbonButton _showCustomDocPropsButton;
        private RibbonButton _setSystemRootButton;
        private RibbonButton _refreshCasePaneButton;
        private RibbonButton _copySampleColumnButton;
        private RibbonButton _resetForDistributionButton;

        /// <summary>
        /// Method: initializes the ribbon.
        /// Args: none.
        /// Returns: none.
        /// Side effects: creates ribbon controls.
        /// </summary>
        public CaseInfoSystemRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary>
        /// Method: builds ribbon controls.
        /// Args: none.
        /// Returns: none.
        /// Side effects: registers the tab, group, and buttons with Excel.
        /// </summary>
        private void InitializeComponent()
        {
            _tab = Factory.CreateRibbonTab();
            _group = Factory.CreateRibbonGroup();
            _showCustomDocPropsButton = Factory.CreateRibbonButton();
            _setSystemRootButton = Factory.CreateRibbonButton();
            _refreshCasePaneButton = Factory.CreateRibbonButton();
            _copySampleColumnButton = Factory.CreateRibbonButton();
            _resetForDistributionButton = Factory.CreateRibbonButton();

            _tab.Label = "\u6848\u4EF6\u60C5\u5831System";
            _tab.Groups.Add(_group);

            _group.Label = "\u4FDD\u5B88";
            _group.Items.Add(_showCustomDocPropsButton);
            _group.Items.Add(_setSystemRootButton);
            _group.Items.Add(_refreshCasePaneButton);
            _group.Items.Add(_copySampleColumnButton);
            _group.Items.Add(_resetForDistributionButton);

            _showCustomDocPropsButton.Label = "DocProp\u4E00\u89A7";
            _showCustomDocPropsButton.ScreenTip = "\u5F53\u8A72\u30D6\u30C3\u30AF\u306E CustomDocumentProperties \u4E00\u89A7\u3092\u8868\u793A\u3057\u307E\u3059\u3002";
            _showCustomDocPropsButton.Click += ShowCustomDocPropsButton_Click;

            _setSystemRootButton.Label = "SystemRoot\u8A2D\u5B9A";
            _setSystemRootButton.ScreenTip = "\u5F53\u8A72\u30D6\u30C3\u30AF\u306E SYSTEM_ROOT \u3092\u624B\u52D5\u8A2D\u5B9A\u3057\u307E\u3059\u3002";
            _setSystemRootButton.Click += SetSystemRootButton_Click;

            _refreshCasePaneButton.Label = "Pane\u66F4\u65B0";
            _refreshCasePaneButton.ScreenTip = "CASE \u30D6\u30C3\u30AF\u306E\u6587\u66F8\u30DC\u30BF\u30F3\u30D1\u30CD\u30EB\u3092\u518D\u63CF\u753B\u3057\u307E\u3059\u3002";
            _refreshCasePaneButton.Click += RefreshCasePaneButton_Click;

            _copySampleColumnButton.Label = "SampleB\u8EE2\u8A18";
            _copySampleColumnButton.ScreenTip = "BASE / CASE \u30D6\u30C3\u30AF\u306E shSample \u30B7\u30FC\u30C8 B\u5217\u3092 shHOME \u3078\u8EE2\u8A18\u3057\u307E\u3059\u3002";
            _copySampleColumnButton.Click += CopySampleColumnButton_Click;

            _resetForDistributionButton.Label = "\u914D\u5E03\u524D\u30EA\u30BB\u30C3\u30C8";
            _resetForDistributionButton.ScreenTip = "Kernel / Base \u30D6\u30C3\u30AF\u306E DocProp \u3092\u914D\u5E03\u7528\u306E\u30AF\u30EA\u30FC\u30F3\u72B6\u614B\u306B\u623B\u3057\u3066\u4FDD\u5B58\u3057\u307E\u3059\u3002";
            _resetForDistributionButton.Click += ResetForDistributionButton_Click;

            Name = "CaseInfoSystemRibbon";
            RibbonType = "Microsoft.Excel.Workbook";
            Tabs.Add(_tab);
        }

        /// <summary>
        /// Method: delegates the custom docprop button click to ThisAddIn.
        /// Args: sender - click source, e - event args.
        /// Returns: none.
        /// Side effects: executes the list display command.
        /// </summary>
        private void ShowCustomDocPropsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn?.ShowActiveWorkbookCustomDocumentProperties();
        }

        /// <summary>
        /// Method: delegates the system root button click to ThisAddIn.
        /// Args: sender - click source, e - event args.
        /// Returns: none.
        /// Side effects: executes the system root update command.
        /// </summary>
        private void SetSystemRootButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn?.SelectAndSaveActiveWorkbookSystemRoot();
        }

        /// <summary>
        /// Method: delegates the case pane refresh button click to ThisAddIn.
        /// Args: sender - click source, e - event args.
        /// Returns: none.
        /// Side effects: refreshes the CASE document button pane when available.
        /// </summary>
        private void RefreshCasePaneButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn?.RefreshActiveWorkbookCaseTaskPane();
        }

        /// <summary>
        /// Method: delegates the sample column copy button click to ThisAddIn.
        /// Args: sender - click source, e - event args.
        /// Returns: none.
        /// Side effects: copies shSample column B values to shHOME.
        /// </summary>
        private void CopySampleColumnButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn?.CopySampleColumnBToHome();
        }

        /// <summary>
        /// Method: delegates the distribution reset button click to ThisAddIn.
        /// Args: sender - click source, e - event args.
        /// Returns: none.
        /// Side effects: executes distribution reset for Kernel/Base workbook.
        /// </summary>
        private void ResetForDistributionButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn?.ResetActiveWorkbookForDistribution();
        }
    }
}
