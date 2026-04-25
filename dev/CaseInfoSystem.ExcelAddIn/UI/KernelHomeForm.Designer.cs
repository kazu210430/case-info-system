using System.ComponentModel;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    internal partial class KernelHomeForm : Form
    {
		private IContainer components = null;

		private Label lblSystemRootTitle;

		private Label lblSystemRootValue;

		private GroupBox grpDateRule;

		private RadioButton optDateNone;

		private RadioButton optDateYY;

		private RadioButton optDateYYYY;

		private GroupBox grpNameRule;

		private RadioButton optNameCustDoc;

		private RadioButton optNameDocCust;

		private RadioButton optNameDoc;

		private GroupBox grpScreenSwitch;

		private Button btnOpenCaseList;

		private Button btnOpenTemplate;

		private Button btnOpenUserData;

		private Button btnExit;

		private Label lblCustomer;

		private TextBox txtCustomer;

		private Button btnCreate;

		private Label lblNewCaseTitlePrefix;

		private Label lblNewCaseTitleFolder;

		private Label lblNewCaseTitleCase;

		private Label lblNewCaseTitleSuffix;

		private Label lblExistingCaseTitlePrefix;

		private Label lblExistingCaseTitleCase;

		private Label lblExistingCaseTitleSuffix;

		private Button btnCreateCaseSingle;

		private Button btnCreateCaseBatch;

		private Panel pnlExistingCaseTree;

		private Label lblExistingTreeDocName;

		private Label lblExistingTreeCaseName;

		private Label lblExistingTreeRootButton;

		private Label label6;

		private Label label2;

		protected override void Dispose (bool disposing)
		{
			if (disposing && components != null) {
				components.Dispose ();
			}
			base.Dispose (disposing);
		}

		private void InitializeComponent ()
		{
            this.lblSystemRootTitle = new System.Windows.Forms.Label();
            this.lblSystemRootValue = new System.Windows.Forms.Label();
            this.grpDateRule = new System.Windows.Forms.GroupBox();
            this.optDateNone = new System.Windows.Forms.RadioButton();
            this.optDateYY = new System.Windows.Forms.RadioButton();
            this.optDateYYYY = new System.Windows.Forms.RadioButton();
            this.grpNameRule = new System.Windows.Forms.GroupBox();
            this.optNameCustDoc = new System.Windows.Forms.RadioButton();
            this.optNameDocCust = new System.Windows.Forms.RadioButton();
            this.optNameDoc = new System.Windows.Forms.RadioButton();
            this.grpScreenSwitch = new System.Windows.Forms.GroupBox();
            this.btnOpenCaseList = new System.Windows.Forms.Button();
            this.btnOpenTemplate = new System.Windows.Forms.Button();
            this.btnOpenUserData = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.lblCustomer = new System.Windows.Forms.Label();
            this.txtCustomer = new System.Windows.Forms.TextBox();
            this.btnCreate = new System.Windows.Forms.Button();
            this.lblNewCaseTitlePrefix = new System.Windows.Forms.Label();
            this.lblNewCaseTitleFolder = new System.Windows.Forms.Label();
            this.lblNewCaseTitleCase = new System.Windows.Forms.Label();
            this.lblNewCaseTitleSuffix = new System.Windows.Forms.Label();
            this.lblExistingCaseTitlePrefix = new System.Windows.Forms.Label();
            this.lblExistingCaseTitleCase = new System.Windows.Forms.Label();
            this.lblExistingCaseTitleSuffix = new System.Windows.Forms.Label();
            this.btnCreateCaseSingle = new System.Windows.Forms.Button();
            this.btnCreateCaseBatch = new System.Windows.Forms.Button();
            this.pnlExistingCaseTree = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.lblExistingTreeDocName = new System.Windows.Forms.Label();
            this.lblExistingTreeCaseName = new System.Windows.Forms.Label();
            this.lblExistingTreeRootButton = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnNewTreeRootButton = new System.Windows.Forms.Button();
            this.lblNewTreeFolderName = new System.Windows.Forms.Label();
            this.lblNewTreeCaseName = new System.Windows.Forms.Label();
            this.lblNewTreeDocName = new System.Windows.Forms.Label();
            this.lblNewTreeRootPath = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.pnlNewCaseTree = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.grpDateRule.SuspendLayout();
            this.grpNameRule.SuspendLayout();
            this.grpScreenSwitch.SuspendLayout();
            this.pnlExistingCaseTree.SuspendLayout();
            this.pnlNewCaseTree.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblSystemRootTitle
            // 
            this.lblSystemRootTitle.AutoSize = true;
            this.lblSystemRootTitle.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblSystemRootTitle.Location = new System.Drawing.Point(24, 14);
            this.lblSystemRootTitle.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblSystemRootTitle.Name = "lblSystemRootTitle";
            this.lblSystemRootTitle.Size = new System.Drawing.Size(100, 17);
            this.lblSystemRootTitle.TabIndex = 0;
            this.lblSystemRootTitle.Text = "システムフォルダ：";
            // 
            // lblSystemRootValue
            // 
            this.lblSystemRootValue.AutoEllipsis = true;
            this.lblSystemRootValue.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblSystemRootValue.ForeColor = System.Drawing.Color.DimGray;
            this.lblSystemRootValue.Location = new System.Drawing.Point(118, 14);
            this.lblSystemRootValue.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblSystemRootValue.Name = "lblSystemRootValue";
            this.lblSystemRootValue.Size = new System.Drawing.Size(520, 35);
            this.lblSystemRootValue.TabIndex = 1;
            this.lblSystemRootValue.Text = "C:\\Users\\kazu2\\Documents\\案件情報System";
            // 
            // grpDateRule
            // 
            this.grpDateRule.BackColor = System.Drawing.Color.WhiteSmoke;
            this.grpDateRule.Controls.Add(this.optDateNone);
            this.grpDateRule.Controls.Add(this.optDateYY);
            this.grpDateRule.Controls.Add(this.optDateYYYY);
            this.grpDateRule.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpDateRule.Location = new System.Drawing.Point(40, 57);
            this.grpDateRule.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.grpDateRule.Name = "grpDateRule";
            this.grpDateRule.Padding = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.grpDateRule.Size = new System.Drawing.Size(123, 127);
            this.grpDateRule.TabIndex = 2;
            this.grpDateRule.TabStop = false;
            this.grpDateRule.Text = "日付の付け方";
            // 
            // optDateNone
            // 
            this.optDateNone.AutoSize = true;
            this.optDateNone.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optDateNone.Location = new System.Drawing.Point(21, 91);
            this.optDateNone.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.optDateNone.Name = "optDateNone";
            this.optDateNone.Size = new System.Drawing.Size(46, 21);
            this.optDateNone.TabIndex = 2;
            this.optDateNone.TabStop = true;
            this.optDateNone.Text = "なし";
            this.optDateNone.UseVisualStyleBackColor = true;
            // 
            // optDateYY
            // 
            this.optDateYY.AutoSize = true;
            this.optDateYY.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optDateYY.Location = new System.Drawing.Point(21, 62);
            this.optDateYY.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.optDateYY.Name = "optDateYY";
            this.optDateYY.Size = new System.Drawing.Size(78, 21);
            this.optDateYY.TabIndex = 1;
            this.optDateYY.TabStop = true;
            this.optDateYY.Text = "yyMMdd";
            this.optDateYY.UseVisualStyleBackColor = true;
            // 
            // optDateYYYY
            // 
            this.optDateYYYY.AutoSize = true;
            this.optDateYYYY.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optDateYYYY.Location = new System.Drawing.Point(21, 33);
            this.optDateYYYY.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.optDateYYYY.Name = "optDateYYYY";
            this.optDateYYYY.Size = new System.Drawing.Size(90, 21);
            this.optDateYYYY.TabIndex = 0;
            this.optDateYYYY.TabStop = true;
            this.optDateYYYY.Text = "yyyyMMdd";
            this.optDateYYYY.UseVisualStyleBackColor = true;
            // 
            // grpNameRule
            // 
            this.grpNameRule.BackColor = System.Drawing.Color.WhiteSmoke;
            this.grpNameRule.Controls.Add(this.optNameCustDoc);
            this.grpNameRule.Controls.Add(this.optNameDocCust);
            this.grpNameRule.Controls.Add(this.optNameDoc);
            this.grpNameRule.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpNameRule.Location = new System.Drawing.Point(183, 57);
            this.grpNameRule.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.grpNameRule.Name = "grpNameRule";
            this.grpNameRule.Padding = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.grpNameRule.Size = new System.Drawing.Size(146, 127);
            this.grpNameRule.TabIndex = 3;
            this.grpNameRule.TabStop = false;
            this.grpNameRule.Text = "ファイル名の表記方法";
            // 
            // optNameCustDoc
            // 
            this.optNameCustDoc.AutoSize = true;
            this.optNameCustDoc.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optNameCustDoc.Location = new System.Drawing.Point(21, 91);
            this.optNameCustDoc.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.optNameCustDoc.Name = "optNameCustDoc";
            this.optNameCustDoc.Size = new System.Drawing.Size(109, 21);
            this.optNameCustDoc.TabIndex = 2;
            this.optNameCustDoc.TabStop = true;
            this.optNameCustDoc.Text = "顧客名_文書名";
            this.optNameCustDoc.UseVisualStyleBackColor = true;
            this.optNameCustDoc.CheckedChanged += new System.EventHandler(this.optNameCustDoc_CheckedChanged);
            // 
            // optNameDocCust
            // 
            this.optNameDocCust.AutoSize = true;
            this.optNameDocCust.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optNameDocCust.Location = new System.Drawing.Point(21, 62);
            this.optNameDocCust.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.optNameDocCust.Name = "optNameDocCust";
            this.optNameDocCust.Size = new System.Drawing.Size(109, 21);
            this.optNameDocCust.TabIndex = 1;
            this.optNameDocCust.TabStop = true;
            this.optNameDocCust.Text = "文書名_顧客名";
            this.optNameDocCust.UseVisualStyleBackColor = true;
            // 
            // optNameDoc
            // 
            this.optNameDoc.AutoSize = true;
            this.optNameDoc.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.optNameDoc.Location = new System.Drawing.Point(21, 33);
            this.optNameDoc.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.optNameDoc.Name = "optNameDoc";
            this.optNameDoc.Size = new System.Drawing.Size(65, 21);
            this.optNameDoc.TabIndex = 0;
            this.optNameDoc.TabStop = true;
            this.optNameDoc.Text = "文書名";
            this.optNameDoc.UseVisualStyleBackColor = true;
            // 
            // grpScreenSwitch
            // 
            this.grpScreenSwitch.BackColor = System.Drawing.Color.WhiteSmoke;
            this.grpScreenSwitch.Controls.Add(this.btnOpenCaseList);
            this.grpScreenSwitch.Controls.Add(this.btnOpenTemplate);
            this.grpScreenSwitch.Controls.Add(this.btnOpenUserData);
            this.grpScreenSwitch.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.grpScreenSwitch.Location = new System.Drawing.Point(522, 57);
            this.grpScreenSwitch.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.grpScreenSwitch.Name = "grpScreenSwitch";
            this.grpScreenSwitch.Padding = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.grpScreenSwitch.Size = new System.Drawing.Size(163, 170);
            this.grpScreenSwitch.TabIndex = 4;
            this.grpScreenSwitch.TabStop = false;
            this.grpScreenSwitch.Text = "画面切替";
            this.grpScreenSwitch.Enter += new System.EventHandler(this.grpScreenSwitch_Enter);
            // 
            // btnOpenCaseList
            // 
            this.btnOpenCaseList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(235)))));
            this.btnOpenCaseList.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnOpenCaseList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenCaseList.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnOpenCaseList.Location = new System.Drawing.Point(21, 121);
            this.btnOpenCaseList.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnOpenCaseList.Name = "btnOpenCaseList";
            this.btnOpenCaseList.Size = new System.Drawing.Size(120, 38);
            this.btnOpenCaseList.TabIndex = 2;
            this.btnOpenCaseList.Text = "案件一覧";
            this.btnOpenCaseList.UseVisualStyleBackColor = false;
            this.btnOpenCaseList.Click += new System.EventHandler(this.btnOpenCaseList_Click);
            // 
            // btnOpenTemplate
            // 
            this.btnOpenTemplate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(235)))));
            this.btnOpenTemplate.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnOpenTemplate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenTemplate.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnOpenTemplate.Location = new System.Drawing.Point(21, 72);
            this.btnOpenTemplate.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnOpenTemplate.Name = "btnOpenTemplate";
            this.btnOpenTemplate.Size = new System.Drawing.Size(120, 38);
            this.btnOpenTemplate.TabIndex = 1;
            this.btnOpenTemplate.Text = "雛形一覧";
            this.btnOpenTemplate.UseVisualStyleBackColor = false;
            // 
            // btnOpenUserData
            // 
            this.btnOpenUserData.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(235)))));
            this.btnOpenUserData.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnOpenUserData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenUserData.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnOpenUserData.Location = new System.Drawing.Point(21, 26);
            this.btnOpenUserData.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnOpenUserData.Name = "btnOpenUserData";
            this.btnOpenUserData.Size = new System.Drawing.Size(120, 38);
            this.btnOpenUserData.TabIndex = 0;
            this.btnOpenUserData.Text = "ユーザー情報";
            this.btnOpenUserData.UseVisualStyleBackColor = false;
            // 
            // btnExit
            // 
            this.btnExit.BackColor = System.Drawing.Color.LightSalmon;
            this.btnExit.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnExit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.btnExit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.OrangeRed;
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExit.Font = new System.Drawing.Font("Yu Gothic UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExit.Location = new System.Drawing.Point(676, 9);
            this.btnExit.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(44, 31);
            this.btnExit.TabIndex = 17;
            this.btnExit.Text = "終 了";
            this.btnExit.UseVisualStyleBackColor = false;
            // 
            // lblCustomer
            // 
            this.lblCustomer.AutoSize = true;
            this.lblCustomer.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblCustomer.Location = new System.Drawing.Point(43, 201);
            this.lblCustomer.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblCustomer.Name = "lblCustomer";
            this.lblCustomer.Size = new System.Drawing.Size(89, 25);
            this.lblCustomer.TabIndex = 5;
            this.lblCustomer.Text = "【顧客名】";
            this.lblCustomer.Click += new System.EventHandler(this.lblCustomer_Click);
            // 
            // txtCustomer
            // 
            this.txtCustomer.AllowDrop = true;
            this.txtCustomer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCustomer.Font = new System.Drawing.Font("BIZ UDPゴシック", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtCustomer.ImeMode = System.Windows.Forms.ImeMode.On;
            this.txtCustomer.Location = new System.Drawing.Point(40, 235);
            this.txtCustomer.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtCustomer.Name = "txtCustomer";
            this.txtCustomer.Size = new System.Drawing.Size(645, 28);
            this.txtCustomer.TabIndex = 6;
            // 
            // btnCreate
            // 
            this.btnCreate.BackColor = System.Drawing.Color.RoyalBlue;
            this.btnCreate.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCreate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCreate.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCreate.FlatAppearance.MouseDownBackColor = System.Drawing.Color.DeepSkyBlue;
            this.btnCreate.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DeepSkyBlue;
            this.btnCreate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreate.Font = new System.Drawing.Font("Yu Gothic UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCreate.ForeColor = System.Drawing.Color.White;
            this.btnCreate.Location = new System.Drawing.Point(543, 271);
            this.btnCreate.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(120, 35);
            this.btnCreate.TabIndex = 7;
            this.btnCreate.Text = "作　 成";
            this.btnCreate.UseVisualStyleBackColor = false;
            // 
            // lblNewCaseTitlePrefix
            // 
            this.lblNewCaseTitlePrefix.AutoSize = true;
            this.lblNewCaseTitlePrefix.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewCaseTitlePrefix.Location = new System.Drawing.Point(50, 275);
            this.lblNewCaseTitlePrefix.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewCaseTitlePrefix.Name = "lblNewCaseTitlePrefix";
            this.lblNewCaseTitlePrefix.Size = new System.Drawing.Size(128, 25);
            this.lblNewCaseTitlePrefix.TabIndex = 8;
            this.lblNewCaseTitlePrefix.Text = "新規の案件 ➡";
            // 
            // lblNewCaseTitleFolder
            // 
            this.lblNewCaseTitleFolder.AutoSize = true;
            this.lblNewCaseTitleFolder.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewCaseTitleFolder.ForeColor = System.Drawing.Color.Green;
            this.lblNewCaseTitleFolder.Location = new System.Drawing.Point(181, 275);
            this.lblNewCaseTitleFolder.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewCaseTitleFolder.Name = "lblNewCaseTitleFolder";
            this.lblNewCaseTitleFolder.Size = new System.Drawing.Size(107, 25);
            this.lblNewCaseTitleFolder.TabIndex = 9;
            this.lblNewCaseTitleFolder.Text = "新規フォルダ";
            // 
            // lblNewCaseTitleCase
            // 
            this.lblNewCaseTitleCase.AutoSize = true;
            this.lblNewCaseTitleCase.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewCaseTitleCase.ForeColor = System.Drawing.Color.Blue;
            this.lblNewCaseTitleCase.Location = new System.Drawing.Point(308, 275);
            this.lblNewCaseTitleCase.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewCaseTitleCase.Name = "lblNewCaseTitleCase";
            this.lblNewCaseTitleCase.Size = new System.Drawing.Size(123, 25);
            this.lblNewCaseTitleCase.TabIndex = 11;
            this.lblNewCaseTitleCase.Text = "案件情報.xlsx";
            this.lblNewCaseTitleCase.Click += new System.EventHandler(this.lblNewCaseTitleCase_Click);
            // 
            // lblNewCaseTitleSuffix
            // 
            this.lblNewCaseTitleSuffix.AutoSize = true;
            this.lblNewCaseTitleSuffix.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewCaseTitleSuffix.Location = new System.Drawing.Point(430, 275);
            this.lblNewCaseTitleSuffix.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewCaseTitleSuffix.Name = "lblNewCaseTitleSuffix";
            this.lblNewCaseTitleSuffix.Size = new System.Drawing.Size(66, 25);
            this.lblNewCaseTitleSuffix.TabIndex = 12;
            this.lblNewCaseTitleSuffix.Text = "の作成";
            this.lblNewCaseTitleSuffix.Click += new System.EventHandler(this.lblNewCaseTitleSuffix_Click);
            // 
            // lblExistingCaseTitlePrefix
            // 
            this.lblExistingCaseTitlePrefix.AutoSize = true;
            this.lblExistingCaseTitlePrefix.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExistingCaseTitlePrefix.Location = new System.Drawing.Point(50, 529);
            this.lblExistingCaseTitlePrefix.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblExistingCaseTitlePrefix.Name = "lblExistingCaseTitlePrefix";
            this.lblExistingCaseTitlePrefix.Size = new System.Drawing.Size(128, 25);
            this.lblExistingCaseTitlePrefix.TabIndex = 13;
            this.lblExistingCaseTitlePrefix.Text = "既存の案件 ➡";
            this.lblExistingCaseTitlePrefix.Click += new System.EventHandler(this.lblExistingCaseTitlePrefix_Click);
            // 
            // lblExistingCaseTitleCase
            // 
            this.lblExistingCaseTitleCase.AutoSize = true;
            this.lblExistingCaseTitleCase.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExistingCaseTitleCase.ForeColor = System.Drawing.Color.Blue;
            this.lblExistingCaseTitleCase.Location = new System.Drawing.Point(181, 529);
            this.lblExistingCaseTitleCase.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblExistingCaseTitleCase.Name = "lblExistingCaseTitleCase";
            this.lblExistingCaseTitleCase.Size = new System.Drawing.Size(123, 25);
            this.lblExistingCaseTitleCase.TabIndex = 14;
            this.lblExistingCaseTitleCase.Text = "案件情報.xlsx";
            // 
            // lblExistingCaseTitleSuffix
            // 
            this.lblExistingCaseTitleSuffix.AutoSize = true;
            this.lblExistingCaseTitleSuffix.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExistingCaseTitleSuffix.Location = new System.Drawing.Point(302, 529);
            this.lblExistingCaseTitleSuffix.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblExistingCaseTitleSuffix.Name = "lblExistingCaseTitleSuffix";
            this.lblExistingCaseTitleSuffix.Size = new System.Drawing.Size(99, 25);
            this.lblExistingCaseTitleSuffix.TabIndex = 15;
            this.lblExistingCaseTitleSuffix.Text = "のみの作成";
            // 
            // btnCreateCaseSingle
            // 
            this.btnCreateCaseSingle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(235)))));
            this.btnCreateCaseSingle.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCreateCaseSingle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateCaseSingle.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCreateCaseSingle.Location = new System.Drawing.Point(411, 525);
            this.btnCreateCaseSingle.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnCreateCaseSingle.Name = "btnCreateCaseSingle";
            this.btnCreateCaseSingle.Size = new System.Drawing.Size(120, 35);
            this.btnCreateCaseSingle.TabIndex = 11;
            this.btnCreateCaseSingle.Text = "単体作成";
            this.btnCreateCaseSingle.UseVisualStyleBackColor = false;
            // 
            // btnCreateCaseBatch
            // 
            this.btnCreateCaseBatch.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(235)))));
            this.btnCreateCaseBatch.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCreateCaseBatch.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateCaseBatch.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnCreateCaseBatch.Location = new System.Drawing.Point(543, 525);
            this.btnCreateCaseBatch.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.btnCreateCaseBatch.Name = "btnCreateCaseBatch";
            this.btnCreateCaseBatch.Size = new System.Drawing.Size(120, 35);
            this.btnCreateCaseBatch.TabIndex = 12;
            this.btnCreateCaseBatch.Text = "連続作成";
            this.btnCreateCaseBatch.UseVisualStyleBackColor = false;
            // 
            // pnlExistingCaseTree
            // 
            this.pnlExistingCaseTree.BackColor = System.Drawing.Color.WhiteSmoke;
            this.pnlExistingCaseTree.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlExistingCaseTree.Controls.Add(this.label2);
            this.pnlExistingCaseTree.Controls.Add(this.lblExistingTreeDocName);
            this.pnlExistingCaseTree.Controls.Add(this.lblExistingTreeCaseName);
            this.pnlExistingCaseTree.Controls.Add(this.lblExistingTreeRootButton);
            this.pnlExistingCaseTree.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.pnlExistingCaseTree.Location = new System.Drawing.Point(40, 568);
            this.pnlExistingCaseTree.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.pnlExistingCaseTree.Name = "pnlExistingCaseTree";
            this.pnlExistingCaseTree.Size = new System.Drawing.Size(645, 156);
            this.pnlExistingCaseTree.TabIndex = 13;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("BIZ UDPゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.Color.Gray;
            this.label2.Location = new System.Drawing.Point(57, 48);
            this.label2.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 92);
            this.label2.TabIndex = 14;
            this.label2.Text = "│\r\n├─\r\n│\r\n├─\r\n│";
            // 
            // lblExistingTreeDocName
            // 
            this.lblExistingTreeDocName.AutoSize = true;
            this.lblExistingTreeDocName.Font = new System.Drawing.Font("BIZ UDゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExistingTreeDocName.ForeColor = System.Drawing.Color.DimGray;
            this.lblExistingTreeDocName.Location = new System.Drawing.Point(96, 101);
            this.lblExistingTreeDocName.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblExistingTreeDocName.Name = "lblExistingTreeDocName";
            this.lblExistingTreeDocName.Size = new System.Drawing.Size(217, 13);
            this.lblExistingTreeDocName.TabIndex = 4;
            this.lblExistingTreeDocName.Text = "20260331_訴状_(例)案件太郎.docx";
            // 
            // lblExistingTreeCaseName
            // 
            this.lblExistingTreeCaseName.AutoSize = true;
            this.lblExistingTreeCaseName.Font = new System.Drawing.Font("BIZ UDゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExistingTreeCaseName.ForeColor = System.Drawing.Color.RoyalBlue;
            this.lblExistingTreeCaseName.Location = new System.Drawing.Point(96, 68);
            this.lblExistingTreeCaseName.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblExistingTreeCaseName.Name = "lblExistingTreeCaseName";
            this.lblExistingTreeCaseName.Size = new System.Drawing.Size(145, 13);
            this.lblExistingTreeCaseName.TabIndex = 2;
            this.lblExistingTreeCaseName.Text = "案件情報_(例)案件太郎";
            this.lblExistingTreeCaseName.Click += new System.EventHandler(this.lblExistingTreeCaseName_Click);
            // 
            // lblExistingTreeRootButton
            // 
            this.lblExistingTreeRootButton.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblExistingTreeRootButton.Font = new System.Drawing.Font("BIZ UDPゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblExistingTreeRootButton.ForeColor = System.Drawing.Color.DimGray;
            this.lblExistingTreeRootButton.Location = new System.Drawing.Point(28, 19);
            this.lblExistingTreeRootButton.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblExistingTreeRootButton.Name = "lblExistingTreeRootButton";
            this.lblExistingTreeRootButton.Size = new System.Drawing.Size(157, 29);
            this.lblExistingTreeRootButton.TabIndex = 0;
            this.lblExistingTreeRootButton.Text = "選択した任意のフォルダ";
            this.lblExistingTreeRootButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Yu Gothic UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(282, 275);
            this.label6.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(31, 25);
            this.label6.TabIndex = 18;
            this.label6.Text = "＆";
            // 
            // btnNewTreeRootButton
            // 
            this.btnNewTreeRootButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(235)))));
            this.btnNewTreeRootButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNewTreeRootButton.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnNewTreeRootButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNewTreeRootButton.Font = new System.Drawing.Font("BIZ UDPゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNewTreeRootButton.ForeColor = System.Drawing.Color.Black;
            this.btnNewTreeRootButton.Location = new System.Drawing.Point(28, 16);
            this.btnNewTreeRootButton.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnNewTreeRootButton.Name = "btnNewTreeRootButton";
            this.btnNewTreeRootButton.Size = new System.Drawing.Size(195, 29);
            this.btnNewTreeRootButton.TabIndex = 0;
            this.btnNewTreeRootButton.Text = "新規ﾌｫﾙﾀﾞの親(保存先)ﾌｫﾙﾀﾞ";
            this.btnNewTreeRootButton.UseVisualStyleBackColor = false;
            // 
            // lblNewTreeFolderName
            // 
            this.lblNewTreeFolderName.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.lblNewTreeFolderName.Font = new System.Drawing.Font("BIZ UDゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewTreeFolderName.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblNewTreeFolderName.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblNewTreeFolderName.Location = new System.Drawing.Point(188, 59);
            this.lblNewTreeFolderName.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewTreeFolderName.Name = "lblNewTreeFolderName";
            this.lblNewTreeFolderName.Size = new System.Drawing.Size(456, 24);
            this.lblNewTreeFolderName.TabIndex = 2;
            this.lblNewTreeFolderName.Text = "20260331_(例)案件太郎";
            this.lblNewTreeFolderName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblNewTreeFolderName.Click += new System.EventHandler(this.lblNewTreeFolderName_Click);
            // 
            // lblNewTreeCaseName
            // 
            this.lblNewTreeCaseName.AutoSize = true;
            this.lblNewTreeCaseName.Font = new System.Drawing.Font("BIZ UDゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewTreeCaseName.ForeColor = System.Drawing.Color.RoyalBlue;
            this.lblNewTreeCaseName.Location = new System.Drawing.Point(170, 103);
            this.lblNewTreeCaseName.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewTreeCaseName.Name = "lblNewTreeCaseName";
            this.lblNewTreeCaseName.Size = new System.Drawing.Size(145, 13);
            this.lblNewTreeCaseName.TabIndex = 4;
            this.lblNewTreeCaseName.Text = "案件情報_(例)案件太郎";
            this.lblNewTreeCaseName.Click += new System.EventHandler(this.lblNewTreeCaseName_Click);
            // 
            // lblNewTreeDocName
            // 
            this.lblNewTreeDocName.AutoSize = true;
            this.lblNewTreeDocName.Font = new System.Drawing.Font("BIZ UDゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewTreeDocName.ForeColor = System.Drawing.Color.DimGray;
            this.lblNewTreeDocName.Location = new System.Drawing.Point(170, 138);
            this.lblNewTreeDocName.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewTreeDocName.Name = "lblNewTreeDocName";
            this.lblNewTreeDocName.Size = new System.Drawing.Size(217, 13);
            this.lblNewTreeDocName.TabIndex = 7;
            this.lblNewTreeDocName.Text = "20260331_訴状_(例)案件太郎.docx";
            // 
            // lblNewTreeRootPath
            // 
            this.lblNewTreeRootPath.AutoEllipsis = true;
            this.lblNewTreeRootPath.Font = new System.Drawing.Font("Yu Gothic UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNewTreeRootPath.ForeColor = System.Drawing.Color.DimGray;
            this.lblNewTreeRootPath.Location = new System.Drawing.Point(228, 21);
            this.lblNewTreeRootPath.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.lblNewTreeRootPath.Name = "lblNewTreeRootPath";
            this.lblNewTreeRootPath.Size = new System.Drawing.Size(410, 42);
            this.lblNewTreeRootPath.TabIndex = 10;
            this.lblNewTreeRootPath.Text = "C:\\Users\\kazu2\\OneDrive\\相談フォルダ";
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("BIZ UDPゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.ForeColor = System.Drawing.Color.Gray;
            this.label4.Location = new System.Drawing.Point(60, 49);
            this.label4.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 45);
            this.label4.TabIndex = 12;
            this.label4.Text = "│\r\n├─\r\n│";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("BIZ UDPゴシック", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.Color.Gray;
            this.label1.Location = new System.Drawing.Point(132, 86);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 75);
            this.label1.TabIndex = 13;
            this.label1.Text = "│\r\n├─\r\n│\r\n├─\r\n│";
            // 
            // pnlNewCaseTree
            // 
            this.pnlNewCaseTree.BackColor = System.Drawing.Color.WhiteSmoke;
            this.pnlNewCaseTree.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlNewCaseTree.Controls.Add(this.label5);
            this.pnlNewCaseTree.Controls.Add(this.label1);
            this.pnlNewCaseTree.Controls.Add(this.label4);
            this.pnlNewCaseTree.Controls.Add(this.lblNewTreeRootPath);
            this.pnlNewCaseTree.Controls.Add(this.lblNewTreeDocName);
            this.pnlNewCaseTree.Controls.Add(this.lblNewTreeCaseName);
            this.pnlNewCaseTree.Controls.Add(this.lblNewTreeFolderName);
            this.pnlNewCaseTree.Controls.Add(this.btnNewTreeRootButton);
            this.pnlNewCaseTree.Location = new System.Drawing.Point(40, 318);
            this.pnlNewCaseTree.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.pnlNewCaseTree.Name = "pnlNewCaseTree";
            this.pnlNewCaseTree.Size = new System.Drawing.Size(645, 190);
            this.pnlNewCaseTree.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label5.Font = new System.Drawing.Font("BIZ UDPゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.ForeColor = System.Drawing.Color.Green;
            this.label5.Location = new System.Drawing.Point(91, 57);
            this.label5.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(94, 29);
            this.label5.TabIndex = 15;
            this.label5.Text = "新規フォルダ";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // KernelHomeForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(731, 769);
            this.Controls.Add(this.pnlExistingCaseTree);
            this.Controls.Add(this.btnCreateCaseBatch);
            this.Controls.Add(this.btnCreateCaseSingle);
            this.Controls.Add(this.lblExistingCaseTitleSuffix);
            this.Controls.Add(this.lblExistingCaseTitleCase);
            this.Controls.Add(this.lblExistingCaseTitlePrefix);
            this.Controls.Add(this.pnlNewCaseTree);
            this.Controls.Add(this.lblNewCaseTitleSuffix);
            this.Controls.Add(this.lblNewCaseTitleCase);
            this.Controls.Add(this.lblNewCaseTitlePrefix);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.txtCustomer);
            this.Controls.Add(this.lblCustomer);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.grpScreenSwitch);
            this.Controls.Add(this.grpNameRule);
            this.Controls.Add(this.grpDateRule);
            this.Controls.Add(this.lblSystemRootValue);
            this.Controls.Add(this.lblSystemRootTitle);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblNewCaseTitleFolder);
            this.Font = new System.Drawing.Font("Yu Gothic UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.Name = "KernelHomeForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "案件情報 作成 HOME";
            this.grpDateRule.ResumeLayout(false);
            this.grpDateRule.PerformLayout();
            this.grpNameRule.ResumeLayout(false);
            this.grpNameRule.PerformLayout();
            this.grpScreenSwitch.ResumeLayout(false);
            this.pnlExistingCaseTree.ResumeLayout(false);
            this.pnlExistingCaseTree.PerformLayout();
            this.pnlNewCaseTree.ResumeLayout(false);
            this.pnlNewCaseTree.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

        private Button btnNewTreeRootButton;
        private Label lblNewTreeFolderName;
        private Label lblNewTreeCaseName;
        private Label lblNewTreeDocName;
        private Label lblNewTreeRootPath;
        private Label label4;
        private Label label1;
        private Panel pnlNewCaseTree;
        private Label label5;
    }
}
