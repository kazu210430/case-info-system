using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.UI
{
	internal partial class KernelHomeForm : Form
	{
		private const string KernelFlickerTracePrefix = "[KernelFlickerTrace]";
		private const string CustomerPlaceholder = "(例)案件太郎";

		private const string PreviewDocumentName = "訴状";

		private const string PreviewCaseWorkbookExtension = ".xlsx";

		private const int SW_SHOWNORMAL = 1;

		private const int GCS_COMPSTR = 8;

		private const int ForegroundRetryIntervalMs = 250;

		private const int ForegroundRetryMaxCount = 8;

		private const int CaseCreationStartMinimizeDelayMs = 2000;

		private static readonly Color NewTreeRootButtonBorderColor = Color.Black;

		private static readonly Color NewTreeRootButtonHoverBorderColor = Color.FromArgb (0, 120, 215);

		private static readonly Color NewTreeRootButtonPressedBorderColor = Color.FromArgb (0, 120, 215);

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly KernelCaseCreationCommandService _kernelCaseCreationCommandService;

		private readonly Logger _logger;

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

		private string _kernelFlickerTraceId = string.Empty;

		private bool _isNewTreeRootButtonHovered;

		private bool _isNewTreeRootButtonPressed;


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

		internal KernelHomeForm (KernelWorkbookService kernelWorkbookService, KernelCaseCreationCommandService kernelCaseCreationCommandService, Logger logger)
			: this ()
		{
			if (kernelWorkbookService == null) {
				throw new ArgumentNullException ("kernelWorkbookService");
			}
			if (kernelCaseCreationCommandService == null) {
				throw new ArgumentNullException ("kernelCaseCreationCommandService");
			}
			if (logger == null) {
				throw new ArgumentNullException ("logger");
			}
			_kernelWorkbookService = kernelWorkbookService;
			_kernelCaseCreationCommandService = kernelCaseCreationCommandService;
			_logger = logger;
			InitializeRuntime ();
		}

		private void InitializeRuntime ()
		{
			base.ShowInTaskbar = true;
			base.WindowState = FormWindowState.Normal;
			WireEvents ();
			ConfigureNewTreeRootButtonAppearance ();
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
			lblNewTreeRootButton.MouseEnter += LblNewTreeRootButton_MouseEnter;
			lblNewTreeRootButton.MouseLeave += LblNewTreeRootButton_MouseLeave;
			lblNewTreeRootButton.MouseDown += LblNewTreeRootButton_MouseDown;
			lblNewTreeRootButton.MouseUp += LblNewTreeRootButton_MouseUp;
			lblNewTreeRootButton.Paint += LblNewTreeRootButton_Paint;
			base.Shown += KernelHomeForm_Shown;
			base.FormClosed += KernelHomeForm_FormClosed;
			ApplyHandCursorToButtons (this);
		}

		private void ConfigureNewTreeRootButtonAppearance ()
		{
			if (lblNewTreeRootButton == null) {
				return;
			}
			lblNewTreeRootButton.BorderStyle = BorderStyle.None;
			lblNewTreeRootButton.FlatStyle = FlatStyle.Standard;
			lblNewTreeRootButton.BackColor = Color.FromArgb (255, 252, 245);
			lblNewTreeRootButton.Cursor = Cursors.Hand;
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
				MessageBox.Show ("シートを開けませんでした。" + ThisAddIn.GetPrimaryTraceLogRelativePath () + " を確認してください。", "案件情報System");
				return;
			}
			_keepBackendSessionOnClose = true;
			_restoreKernelWorkbookOnClose = true;
			_sheetNavigationHandled = true;
			StopForegroundRetry ();
			Hide ();
			if (!Globals.ThisAddIn.ShowKernelSheetAndRefreshPane (codeName, "KernelHomeForm.OpenSheet")) {
				MessageBox.Show ("シートを開けませんでした。" + ThisAddIn.GetPrimaryTraceLogRelativePath () + " を確認してください。", "案件情報System");
			} else {
				if (Globals.ThisAddIn != null) {
					Globals.ThisAddIn.ScheduleWorkbookTaskPaneRefresh (_kernelWorkbookService.GetOpenKernelWorkbook (), "KernelHomeForm.OpenSheet.PostClose");
				}
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
				RequestEnsureHomeDisplayHidden ("KernelHomeForm.RestoreHomeToForegroundAfterCaseCreation.BeforeShow");
				if (!Visible) {
					Show ();
				}
				base.WindowState = FormWindowState.Normal;
				ForceBringToFront ("KernelHomeForm.RestoreHomeToForegroundAfterCaseCreation");
				BeginForegroundRetry ();
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
					EnsureKernelFlickerTraceContext ("KernelHomeForm.Shown");
					LogKernelFlickerTrace (
						"source=KernelHomeForm action=shown-begin visible="
						+ base.Visible.ToString ()
						+ ", retryActive="
						+ (_foregroundRetryTimer != null).ToString ()
						+ ", formWindowState="
						+ base.WindowState.ToString ());
					ForceBringToFront ("KernelHomeForm.Shown");
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
			EnsureKernelFlickerTraceContext ("KernelHomeForm.BeginForegroundRetry");
			LogKernelFlickerTrace (
				"source=KernelHomeForm action=foreground-retry-begin intervalMs="
				+ ForegroundRetryIntervalMs.ToString ()
				+ ", maxCount="
				+ ForegroundRetryMaxCount.ToString ()
				+ ", currentRetryCount="
				+ _foregroundRetryCount.ToString ()
				+ ", visible="
				+ base.Visible.ToString ());
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
				EnsureKernelFlickerTraceContext ("KernelHomeForm.StopForegroundRetry");
				LogKernelFlickerTrace (
					"source=KernelHomeForm action=foreground-retry-stop currentRetryCount="
					+ _foregroundRetryCount.ToString ()
					+ ", visible="
					+ base.Visible.ToString ());
				_foregroundRetryTimer.Stop ();
				_foregroundRetryTimer.Tick -= ForegroundRetryTimer_Tick;
				_foregroundRetryTimer.Dispose ();
				_foregroundRetryTimer = null;
			}
		}

		private void ForegroundRetryTimer_Tick (object sender, EventArgs e)
		{
			EnsureKernelFlickerTraceContext ("KernelHomeForm.ForegroundRetryTimer_Tick");
			if (base.IsDisposed || !base.Visible) {
				LogKernelFlickerTrace (
					"source=KernelHomeForm action=foreground-retry-tick-skip reason=form-not-available visible="
					+ base.Visible.ToString ()
					+ ", isDisposed="
					+ base.IsDisposed.ToString ()
					+ ", retryCount="
					+ _foregroundRetryCount.ToString ());
				StopForegroundRetry ();
				return;
			}
			_foregroundRetryCount++;
			LogKernelFlickerTrace (
				"source=KernelHomeForm action=foreground-retry-tick retryCount="
				+ _foregroundRetryCount.ToString ()
				+ ", formWindowState="
				+ base.WindowState.ToString ()
				+ ", activeForm="
				+ (Form.ActiveForm == null ? string.Empty : Form.ActiveForm.Name ?? string.Empty));
			ForceBringToFront ("KernelHomeForm.ForegroundRetryTimer_Tick");
			RequestEnsureHomeDisplayHidden ("KernelHomeForm.ForegroundRetryTimer_Tick.PostForceBringToFront");
			if (_foregroundRetryCount >= 8) {
				StopForegroundRetry ();
			}
		}

		private void ForceBringToFront (string triggerSource)
		{
			EnsureKernelFlickerTraceContext (triggerSource);
			LogKernelFlickerTrace (
				"source=KernelHomeForm action=force-bring-to-front-enter trigger="
				+ (triggerSource ?? string.Empty)
				+ ", retryCount="
				+ _foregroundRetryCount.ToString ()
				+ ", visible="
				+ base.Visible.ToString ()
				+ ", topMost="
				+ base.TopMost.ToString ());
			RequestEnsureHomeDisplayHidden ((triggerSource ?? string.Empty) + ".PreBringToFront");
			base.TopMost = true;
			ShowWindow (base.Handle, 1);
			Activate ();
			BringToFront ();
			SetForegroundWindow (base.Handle);
			base.TopMost = false;
			LogKernelFlickerTrace (
				"source=KernelHomeForm action=force-bring-to-front-end trigger="
				+ (triggerSource ?? string.Empty)
				+ ", retryCount="
				+ _foregroundRetryCount.ToString ()
				+ ", visible="
				+ base.Visible.ToString ()
				+ ", topMost="
				+ base.TopMost.ToString ());
		}

		private void RequestEnsureHomeDisplayHidden (string triggerSource)
		{
			EnsureKernelFlickerTraceContext (triggerSource);
			string requestReason = "source="
				+ (triggerSource ?? string.Empty)
				+ ",foregroundRetryCount="
				+ _foregroundRetryCount.ToString ()
				+ ",formVisible="
				+ base.Visible.ToString ()
				+ ",formWindowState="
				+ base.WindowState.ToString ();
			LogKernelFlickerTrace (
				"source=KernelHomeForm action=ensure-home-display-hidden-request "
				+ requestReason);
			_kernelWorkbookService.EnsureHomeDisplayHidden (requestReason);
		}

		private void EnsureKernelFlickerTraceContext (string triggerSource)
		{
			string currentTraceId = KernelFlickerTraceContext.CurrentTraceId;
			if (!string.IsNullOrWhiteSpace (currentTraceId)) {
				_kernelFlickerTraceId = currentTraceId;
				return;
			}
			if (string.IsNullOrWhiteSpace (_kernelFlickerTraceId)) {
				return;
			}
			KernelFlickerTraceContext.SetCurrentTrace (_kernelFlickerTraceId);
			LogKernelFlickerTrace (
				"source=KernelHomeForm action=trace-restore trigger="
				+ (triggerSource ?? string.Empty)
				+ ", restoredTraceId="
				+ _kernelFlickerTraceId);
		}

		private void LogKernelFlickerTrace (string detail)
		{
			_logger.Info (KernelFlickerTracePrefix + " " + (detail ?? string.Empty));
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

		private void LblNewTreeRootButton_MouseEnter (object sender, EventArgs e)
		{
			_isNewTreeRootButtonHovered = true;
			lblNewTreeRootButton.Invalidate ();
		}

		private void LblNewTreeRootButton_MouseLeave (object sender, EventArgs e)
		{
			_isNewTreeRootButtonHovered = false;
			_isNewTreeRootButtonPressed = false;
			lblNewTreeRootButton.Invalidate ();
		}

		private void LblNewTreeRootButton_MouseDown (object sender, MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Left) {
				return;
			}
			_isNewTreeRootButtonPressed = true;
			lblNewTreeRootButton.Invalidate ();
		}

		private void LblNewTreeRootButton_MouseUp (object sender, MouseEventArgs e)
		{
			_isNewTreeRootButtonPressed = false;
			lblNewTreeRootButton.Invalidate ();
		}

		private void LblNewTreeRootButton_Paint (object sender, PaintEventArgs e)
		{
			Rectangle rectangle = lblNewTreeRootButton.ClientRectangle;
			if (rectangle.Width <= 2 || rectangle.Height <= 2) {
				return;
			}
			Color color = GetNewTreeRootButtonBorderColor ();
			rectangle.Width--;
			rectangle.Height--;
			using (Pen pen = new Pen (color)) {
				e.Graphics.DrawRectangle (pen, rectangle);
			}
		}

		private Color GetNewTreeRootButtonBorderColor ()
		{
			if (_isNewTreeRootButtonPressed) {
				return NewTreeRootButtonPressedBorderColor;
			}
			if (_isNewTreeRootButtonHovered) {
				return NewTreeRootButtonHoverBorderColor;
			}
			return NewTreeRootButtonBorderColor;
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

        private void lblNewTreeFolderName_Click(object sender, EventArgs e)
        {

        }
    }
}
