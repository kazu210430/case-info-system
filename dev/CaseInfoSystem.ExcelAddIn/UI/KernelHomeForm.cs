using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal partial class KernelHomeForm : Form
	{
		private const string CustomerPlaceholder = "(例)案件太郎";

		private const string PreviewDocumentName = "訴状";

		private const string PreviewCaseWorkbookExtension = ".xlsx";

		private const int SW_SHOWNORMAL = 1;

		private const int GCS_COMPSTR = 8;

		private const int ForegroundRetryIntervalMs = 250;

		private const int ForegroundRetryMaxCount = 8;

		private const int CaseCreationStartMinimizeDelayMs = 2000;

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly KernelCaseCreationCommandService _kernelCaseCreationCommandService;

		private bool _isInitializing;

		private bool _keepBackendSessionOnClose;

		private bool _restoreKernelWorkbookOnClose;

		private bool _saveKernelWorkbookOnClose;

		private Timer _foregroundRetryTimer;

		private Timer _caseCreationStartMinimizeTimer;

		private int _foregroundRetryCount;

		private bool _sheetNavigationHandled;

		private bool _caseCreationStartMinimizeActive;

		private FormWindowState _caseCreationStartMinimizePreviousState = FormWindowState.Normal;

		private IDisposable _kernelCaseCreationFlowScope;

		private bool _isClosingBySession;


		[DllImport ("user32.dll")]
		private static extern bool SetForegroundWindow (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern bool ShowWindow (IntPtr hWnd, int nCmdShow);

		[DllImport ("imm32.dll")]
		private static extern IntPtr ImmGetContext (IntPtr hWnd);

		[DllImport ("imm32.dll")]
		private static extern bool ImmReleaseContext (IntPtr hWnd, IntPtr hIMC);

		[DllImport ("imm32.dll")]
		private static extern int ImmGetCompositionString (IntPtr hIMC, int dwIndex, byte[] lpBuf, int dwBufLen);

		public KernelHomeForm ()
		{
			InitializeComponent ();
		}

		internal KernelHomeForm (KernelWorkbookService kernelWorkbookService, KernelCaseCreationCommandService kernelCaseCreationCommandService)
			: this ()
		{
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_kernelCaseCreationCommandService = kernelCaseCreationCommandService ?? throw new ArgumentNullException ("kernelCaseCreationCommandService");
			InitializeRuntime ();
		}

		private void InitializeRuntime ()
		{
			base.ShowInTaskbar = true;
			base.WindowState = FormWindowState.Normal;
			WireEvents ();
			LoadSettings ();
			SetCustomerPlaceholder ();
			RefreshPreview ();
		}

		private void WireEvents ()
		{
			txtCustomer.TextChanged += delegate {
				RefreshPreview ();
			};
			txtCustomer.Enter += TxtCustomer_Enter;
			txtCustomer.Leave += TxtCustomer_Leave;
			txtCustomer.KeyDown += TxtCustomer_KeyDown;
			optDateYYYY.CheckedChanged += delegate {
				HandleNameRuleAChanged ("YYYY", optDateYYYY.Checked);
			};
			optDateYY.CheckedChanged += delegate {
				HandleNameRuleAChanged ("YY", optDateYY.Checked);
			};
			optDateNone.CheckedChanged += delegate {
				HandleNameRuleAChanged ("NONE", optDateNone.Checked);
			};
			optNameDoc.CheckedChanged += delegate {
				HandleNameRuleBChanged ("DOC", optNameDoc.Checked);
			};
			optNameDocCust.CheckedChanged += delegate {
				HandleNameRuleBChanged ("DOC_CUST", optNameDocCust.Checked);
			};
			optNameCustDoc.CheckedChanged += delegate {
				HandleNameRuleBChanged ("CUST_DOC", optNameCustDoc.Checked);
			};
			btnCreate.Click += BtnCreate_Click;
			btnExit.Click += delegate {
				CloseHomeSession ();
			};
			btnOpenUserData.Click += delegate {
				OpenSheet ("shUserData", "ユーザー情報");
			};
			btnOpenTemplate.Click += delegate {
				OpenSheet ("shMasterList", "雛形登録");
			};
			btnOpenCaseList.Click += delegate {
				OpenSheet ("shCaseList", "案件一覧");
			};
			btnCreateCaseSingle.Click += delegate {
				RunCreateCase (showCase: true);
			};
			btnCreateCaseBatch.Click += delegate {
				RunCreateCase (showCase: false);
			};
			lblNewTreeRootButton.Click += delegate {
				SelectDefaultRoot ();
			};
			base.Shown += KernelHomeForm_Shown;
			base.FormClosed += KernelHomeForm_FormClosed;
			ApplyHandCursorToButtons (this);
		}

		private void LoadSettings ()
		{
			_isInitializing = true;
			try {
				KernelSettingsState kernelSettingsState = _kernelWorkbookService.LoadSettings ();
				lblSystemRootValue.Text = (string.IsNullOrWhiteSpace (kernelSettingsState.SystemRoot) ? "（未設定）" : kernelSettingsState.SystemRoot);
				ApplyDefaultRootDisplay (kernelSettingsState.DefaultRoot);
				optDateYYYY.Checked = kernelSettingsState.NameRuleA == "YYYY";
				optDateYY.Checked = kernelSettingsState.NameRuleA == "YY";
				optDateNone.Checked = kernelSettingsState.NameRuleA == "NONE";
				optNameDoc.Checked = kernelSettingsState.NameRuleB == "DOC";
				optNameDocCust.Checked = kernelSettingsState.NameRuleB == "DOC_CUST";
				optNameCustDoc.Checked = kernelSettingsState.NameRuleB == "CUST_DOC";
			} finally {
				_isInitializing = false;
			}
		}

		internal void ReloadSettings ()
		{
			LoadSettings ();
			RefreshPreview ();
		}

		private void RefreshPreview ()
		{
			string customerDisplayName = GetCustomerDisplayName ();
			string selectedNameRuleA = GetSelectedNameRuleA ();
			string selectedNameRuleB = GetSelectedNameRuleB ();
			string text = KernelNamingService.BuildDocumentName (selectedNameRuleA, selectedNameRuleB, "訴状", customerDisplayName, DateTime.Today) + ".docx";
			string text2 = EnsureCaseWorkbookExtension (KernelNamingService.BuildCaseBookName (customerDisplayName, ".xlsx"));
			lblNewTreeFolderName.Text = KernelNamingService.BuildFolderName (selectedNameRuleA, customerDisplayName, DateTime.Today);
			lblNewTreeCaseName.Text = text2;
			lblNewTreeDocName.Text = text;
			lblExistingTreeCaseName.Text = text2;
			lblExistingTreeDocName.Text = text;
		}

		private static string EnsureCaseWorkbookExtension (string caseBookName)
		{
			string text = (caseBookName ?? string.Empty).Trim ();
			if (text.EndsWith (".xlsx", StringComparison.OrdinalIgnoreCase)) {
				return text;
			}
			return text + ".xlsx";
		}

		private string GetSelectedNameRuleA ()
		{
			if (optDateYYYY.Checked) {
				return "YYYY";
			}
			if (optDateYY.Checked) {
				return "YY";
			}
			return "NONE";
		}

		private string GetSelectedNameRuleB ()
		{
			if (optNameDocCust.Checked) {
				return "DOC_CUST";
			}
			if (optNameCustDoc.Checked) {
				return "CUST_DOC";
			}
			return "DOC";
		}

		private void BtnCreate_Click (object sender, EventArgs e)
		{
			string actualCustomerName = GetActualCustomerName ();
			if (string.IsNullOrWhiteSpace (actualCustomerName)) {
				MessageBox.Show ("顧客名を入力してください。", "案件情報System");
				return;
			}
			PrepareForCaseCreationStart ();
			BeginKernelCaseCreationFlow ("KernelHomeForm.BtnCreate");
			try {
				KernelCaseCreationResult result = _kernelCaseCreationCommandService.ExecuteNewCaseDefault (actualCustomerName);
				HandleCaseCreationResult (result);
			} catch {
				EndKernelCaseCreationFlow ("KernelHomeForm.BtnCreate.Exception");
				throw;
			}
		}

		private void OpenSheet (string codeName, string featureName)
		{
			if (string.IsNullOrWhiteSpace (codeName)) {
				return;
			}
			if (Globals.ThisAddIn == null) {
				MessageBox.Show ("シートを開けませんでした。logs\\KernelHomeAddin_trace.log を確認してください。", "案件情報System");
				return;
			}
			_keepBackendSessionOnClose = true;
			_restoreKernelWorkbookOnClose = true;
			_sheetNavigationHandled = true;
			StopForegroundRetry ();
			Hide ();
			_kernelWorkbookService.CompleteHomeNavigation (showExcel: true);
			if (!Globals.ThisAddIn.ShowKernelSheetAndRefreshPaneFromHome (codeName, "KernelHomeForm.OpenSheet")) {
				MessageBox.Show ("シートを開けませんでした。logs\\KernelHomeAddin_trace.log を確認してください。", "案件情報System");
			} else {
				Globals.ThisAddIn?.ScheduleWorkbookTaskPaneRefresh (_kernelWorkbookService.GetOpenKernelWorkbook (), "KernelHomeForm.OpenSheet.PostClose");
			}
			Close ();
		}

		private void RunCreateCase (bool showCase)
		{
			string actualCustomerName = GetActualCustomerName ();
			BeginKernelCaseCreationFlow (showCase ? "KernelHomeForm.RunCreateCase.Single" : "KernelHomeForm.RunCreateCase.Batch");
			try {
				KernelCaseCreationResult result = (showCase ? _kernelCaseCreationCommandService.ExecuteCreateCaseSingle (actualCustomerName) : _kernelCaseCreationCommandService.ExecuteCreateCaseBatch (actualCustomerName));
				HandleCaseCreationResult (result);
			} catch {
				EndKernelCaseCreationFlow ("KernelHomeForm.RunCreateCase.Exception");
				throw;
			}
		}

		private void HandleCaseCreationResult (KernelCaseCreationResult result)
		{
			if (result == null) {
				EndKernelCaseCreationFlow ("KernelHomeForm.HandleCaseCreationResult.ResultNull");
				RestoreAfterCaseCreationStartIfNeeded ();
				MessageBox.Show ("案件作成結果を取得できませんでした。", "案件情報System");
				return;
			}
			if (!string.IsNullOrWhiteSpace (result.UserMessage)) {
				RestoreAfterCaseCreationStartIfNeeded ();
				MessageBox.Show (result.UserMessage, "案件情報System");
			}
			if (!result.Success) {
				EndKernelCaseCreationFlow ("KernelHomeForm.HandleCaseCreationResult.Failure");
				RestoreAfterCaseCreationStartIfNeeded ();
			} else if (result.ShouldCloseKernelHome) {
				CloseKernelAfterCaseCreation ();
			} else {
				EndKernelCaseCreationFlow ("KernelHomeForm.HandleCaseCreationResult.SuccessContinue");
				RestoreAfterCaseCreationStartIfNeeded ();
				RestoreHomeToForegroundAfterCaseCreation (result);
			}
		}

		private void RestoreHomeToForegroundAfterCaseCreation (KernelCaseCreationResult result)
		{
			if (result != null && result.Mode == KernelCaseCreationMode.CreateCaseBatch && !base.IsDisposed) {
				Show ();
				base.WindowState = FormWindowState.Normal;
				Activate ();
				BringToFront ();
				ForceBringToFront ();
				PrepareCustomerInputForNextBatchCreate ();
			}
		}

		private void PrepareCustomerInputForNextBatchCreate ()
		{
			if (!base.IsDisposed) {
				txtCustomer.ForeColor = Color.Black;
				txtCustomer.Text = string.Empty;
				txtCustomer.ImeMode = ImeMode.On;
				txtCustomer.Focus ();
				txtCustomer.SelectionStart = 0;
				txtCustomer.SelectionLength = 0;
			}
		}

		private void PrepareForCaseCreationStart ()
		{
			if (!_caseCreationStartMinimizeActive) {
				_caseCreationStartMinimizePreviousState = base.WindowState;
				_caseCreationStartMinimizeActive = true;
				EnsureCaseCreationStartMinimizeTimer ();
				_caseCreationStartMinimizeTimer.Interval = 2000;
				_caseCreationStartMinimizeTimer.Stop ();
				_caseCreationStartMinimizeTimer.Start ();
			}
		}

		private void RestoreAfterCaseCreationStartIfNeeded ()
		{
			if (_caseCreationStartMinimizeActive && !base.IsDisposed) {
				if (_caseCreationStartMinimizeTimer != null) {
					_caseCreationStartMinimizeTimer.Stop ();
				}
				base.WindowState = _caseCreationStartMinimizePreviousState;
				Activate ();
				_caseCreationStartMinimizeActive = false;
				_caseCreationStartMinimizePreviousState = FormWindowState.Normal;
			}
		}

		private void EnsureCaseCreationStartMinimizeTimer ()
		{
			if (_caseCreationStartMinimizeTimer == null) {
				_caseCreationStartMinimizeTimer = new Timer ();
				_caseCreationStartMinimizeTimer.Tick += CaseCreationStartMinimizeTimer_Tick;
			}
		}

		private void CaseCreationStartMinimizeTimer_Tick (object sender, EventArgs e)
		{
			if (_caseCreationStartMinimizeTimer != null) {
				_caseCreationStartMinimizeTimer.Stop ();
			}
			if (_caseCreationStartMinimizeActive && !base.IsDisposed) {
				base.WindowState = FormWindowState.Minimized;
			}
		}

		private void SelectDefaultRoot ()
		{
			string text = _kernelWorkbookService.SelectAndSaveDefaultRoot ();
			if (!string.IsNullOrWhiteSpace (text)) {
				ApplyDefaultRootDisplay (text);
			}
		}

		private void ApplyDefaultRootDisplay (string defaultRoot)
		{
			if (string.IsNullOrWhiteSpace (defaultRoot)) {
				lblNewTreeRootPath.Text = "（未設定）";
				lblNewTreeRootPath.ForeColor = Color.Red;
			} else {
				lblNewTreeRootPath.Text = defaultRoot;
				lblNewTreeRootPath.ForeColor = Color.DimGray;
			}
		}

		private void CloseHomeSession ()
		{
			_isClosingBySession = true;
			Close ();
		}

		internal void PrepareForExistingCaseOpenClose ()
		{
			_isClosingBySession = true;
			_sheetNavigationHandled = false;
			_keepBackendSessionOnClose = true;
			_restoreKernelWorkbookOnClose = false;
			_saveKernelWorkbookOnClose = true;
		}

		private void CloseKernelAfterCaseCreation ()
		{
			_isClosingBySession = true;
			_sheetNavigationHandled = false;
			_keepBackendSessionOnClose = true;
			_restoreKernelWorkbookOnClose = false;
			_saveKernelWorkbookOnClose = true;
			StopForegroundRetry ();
			Hide ();
			Close ();
		}

		private void KernelHomeForm_Shown (object sender, EventArgs e)
		{
			BeginInvoke ((MethodInvoker)delegate {
				if (!base.IsDisposed) {
					ForceBringToFront ();
					BeginForegroundRetry ();
				}
			});
		}

		private void KernelHomeForm_FormClosed (object sender, FormClosedEventArgs e)
		{
			try {
				if (_kernelWorkbookService == null) {
					return;
				}
				StopForegroundRetry ();
				if (!_isClosingBySession) {
					_isClosingBySession = true;
				}
				if (_sheetNavigationHandled) {
					return;
				}
				if (_keepBackendSessionOnClose) {
					if (_saveKernelWorkbookOnClose) {
						_kernelWorkbookService.CloseHomeSessionSavingKernel ();
					} else {
						_kernelWorkbookService.CompleteHomeNavigation (_restoreKernelWorkbookOnClose);
					}
				} else {
					_kernelWorkbookService.CloseHomeSession ();
				}
			} finally {
				EndKernelCaseCreationFlow ("KernelHomeForm.FormClosed");
			}
		}

		private void BeginKernelCaseCreationFlow (string reason)
		{
			if (_kernelCaseCreationFlowScope == null && Globals.ThisAddIn != null && Globals.ThisAddIn.KernelCaseInteractionState != null) {
				_kernelCaseCreationFlowScope = Globals.ThisAddIn.KernelCaseInteractionState.BeginKernelCaseCreationFlow (reason);
			}
		}

		private void EndKernelCaseCreationFlow (string reason)
		{
			if (_kernelCaseCreationFlowScope == null) {
				return;
			}
			try {
				_kernelCaseCreationFlowScope.Dispose ();
			} finally {
				_kernelCaseCreationFlowScope = null;
			}
		}

		private void BeginForegroundRetry ()
		{
			StopForegroundRetry ();
			_foregroundRetryCount = 0;
			_foregroundRetryTimer = new Timer ();
			_foregroundRetryTimer.Interval = 250;
			_foregroundRetryTimer.Tick += ForegroundRetryTimer_Tick;
			_foregroundRetryTimer.Start ();
		}

		private void StopForegroundRetry ()
		{
			if (_foregroundRetryTimer != null) {
				_foregroundRetryTimer.Stop ();
				_foregroundRetryTimer.Tick -= ForegroundRetryTimer_Tick;
				_foregroundRetryTimer.Dispose ();
				_foregroundRetryTimer = null;
			}
		}

		private void ForegroundRetryTimer_Tick (object sender, EventArgs e)
		{
			if (base.IsDisposed || !base.Visible) {
				StopForegroundRetry ();
				return;
			}
			_foregroundRetryCount++;
			ForceBringToFront ();
			_kernelWorkbookService.EnsureHomeDisplayHidden ();
			if (_foregroundRetryCount >= 8) {
				StopForegroundRetry ();
			}
		}

		private void ForceBringToFront ()
		{
			_kernelWorkbookService.EnsureHomeDisplayHidden ();
			base.TopMost = true;
			ShowWindow (base.Handle, 1);
			Activate ();
			BringToFront ();
			SetForegroundWindow (base.Handle);
			base.TopMost = false;
		}

		private void HandleNameRuleAChanged (string ruleA, bool isChecked)
		{
			if (isChecked) {
				if (!_isInitializing) {
					_kernelWorkbookService.SaveNameRuleA (ruleA);
				}
				RefreshPreview ();
			}
		}

		private void HandleNameRuleBChanged (string ruleB, bool isChecked)
		{
			if (isChecked) {
				if (!_isInitializing) {
					_kernelWorkbookService.SaveNameRuleB (ruleB);
				}
				RefreshPreview ();
			}
		}

		private string GetCustomerDisplayName ()
		{
			string actualCustomerName = GetActualCustomerName ();
			return string.IsNullOrWhiteSpace (actualCustomerName) ? "(例)案件太郎" : actualCustomerName;
		}

		private string GetActualCustomerName ()
		{
			string text = txtCustomer.Text.Trim ();
			return (text == "(例)案件太郎") ? string.Empty : text;
		}

		private void TxtCustomer_Enter (object sender, EventArgs e)
		{
			txtCustomer.ImeMode = ImeMode.On;
			if (!(txtCustomer.Text != "(例)案件太郎")) {
				txtCustomer.Text = string.Empty;
				txtCustomer.ForeColor = Color.Black;
			}
		}

		private void TxtCustomer_Leave (object sender, EventArgs e)
		{
			if (string.IsNullOrWhiteSpace (txtCustomer.Text)) {
				SetCustomerPlaceholder ();
			}
		}

		private void SetCustomerPlaceholder ()
		{
			txtCustomer.Text = "(例)案件太郎";
			txtCustomer.ForeColor = Color.Silver;
		}

		private void TxtCustomer_KeyDown (object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Return && !IsImeComposing (txtCustomer)) {
				e.Handled = true;
				e.SuppressKeyPress = true;
				BtnCreate_Click (btnCreate, EventArgs.Empty);
			}
		}

		private void ApplyHandCursorToButtons (Control parent)
		{
			foreach (Control control in parent.Controls) {
				if (control is ButtonBase) {
					control.Cursor = Cursors.Hand;
				}
				if (control.HasChildren) {
					ApplyHandCursorToButtons (control);
				}
			}
		}

		private static bool IsImeComposing (Control control)
		{
			IntPtr intPtr = ImmGetContext (control.Handle);
			if (intPtr == IntPtr.Zero) {
				return false;
			}
			try {
				return ImmGetCompositionString (intPtr, 8, null, 0) > 0;
			} finally {
				ImmReleaseContext (control.Handle, intPtr);
			}
		}

		private void lblCustomer_Click (object sender, EventArgs e)
		{
		}

		private void lblNewCaseTitleCase_Click (object sender, EventArgs e)
		{
		}

		private void lblNewTreeCaseName_Click (object sender, EventArgs e)
		{
		}

		private void label4_Click (object sender, EventArgs e)
		{
		}

		private void lblExistingTreeCaseName_Click (object sender, EventArgs e)
		{
		}

		private void btnOpenCaseList_Click (object sender, EventArgs e)
		{
		}

		private void optNameCustDoc_CheckedChanged (object sender, EventArgs e)
		{
		}

		private void lblNewCaseTitleSuffix_Click (object sender, EventArgs e)
		{
		}

		private void lblExistingCaseTitlePrefix_Click (object sender, EventArgs e)
		{
		}


        private void grpScreenSwitch_Enter(object sender, EventArgs e)
        {

        }
    }
}
