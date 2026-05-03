using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal interface IAccountingSetReadyShowBridge
	{
		void ShowWorkbookTaskPaneWhenReady (Workbook workbook, string reason);
	}

	internal sealed class ThisAddInAccountingSetReadyShowBridge : IAccountingSetReadyShowBridge
	{
		private readonly ThisAddIn _addIn;

		internal ThisAddInAccountingSetReadyShowBridge (ThisAddIn addIn)
		{
			_addIn = addIn ?? throw new ArgumentNullException ("addIn");
		}

		public void ShowWorkbookTaskPaneWhenReady (Workbook workbook, string reason)
		{
			_addIn.ShowWorkbookTaskPaneWhenReady (workbook, reason);
		}
	}

	internal sealed class AccountingSetCreateService
	{
		private const string CustomerNameKey = "\u9867\u5BA2_\u540D\u524D";

		private const string CustomerHonorificKey = "\u9867\u5BA2_\u656C\u79F0";

		private const string LawyerKey = "\u5F53\u65B9_\u5F01\u8B77\u58EB";

		private const string SystemRootPropertyName = "SYSTEM_ROOT";

		private const string NameRuleAPropertyName = "NAME_RULE_A";

		private const string NameRuleBPropertyName = "NAME_RULE_B";

		private readonly ExcelInteropService _excelInteropService;

		private readonly CaseContextFactory _caseContextFactory;

		private readonly DocumentOutputService _documentOutputService;

		private readonly AccountingSetNamingService _accountingSetNamingService;

		private readonly AccountingTemplateResolver _accountingTemplateResolver;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly TransientPaneSuppressionService _transientPaneSuppressionService;

		private readonly AccountingSetPresentationWaitService _accountingSetPresentationWaitService;

		private readonly IAccountingSetReadyShowBridge _accountingSetReadyShowBridge;

		private readonly Logger _logger;

		internal AccountingSetCreateService (ExcelInteropService excelInteropService, CaseContextFactory caseContextFactory, DocumentOutputService documentOutputService, AccountingSetNamingService accountingSetNamingService, AccountingTemplateResolver accountingTemplateResolver, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, TransientPaneSuppressionService transientPaneSuppressionService, AccountingSetPresentationWaitService accountingSetPresentationWaitService, Logger logger)
			: this (excelInteropService, caseContextFactory, documentOutputService, accountingSetNamingService, accountingTemplateResolver, accountingWorkbookService, pathCompatibilityService, transientPaneSuppressionService, accountingSetPresentationWaitService, new ThisAddInAccountingSetReadyShowBridge (Globals.ThisAddIn), logger)
		{
		}

		internal AccountingSetCreateService (ExcelInteropService excelInteropService, CaseContextFactory caseContextFactory, DocumentOutputService documentOutputService, AccountingSetNamingService accountingSetNamingService, AccountingTemplateResolver accountingTemplateResolver, AccountingWorkbookService accountingWorkbookService, PathCompatibilityService pathCompatibilityService, TransientPaneSuppressionService transientPaneSuppressionService, AccountingSetPresentationWaitService accountingSetPresentationWaitService, IAccountingSetReadyShowBridge accountingSetReadyShowBridge, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_caseContextFactory = caseContextFactory ?? throw new ArgumentNullException ("caseContextFactory");
			_documentOutputService = documentOutputService ?? throw new ArgumentNullException ("documentOutputService");
			_accountingSetNamingService = accountingSetNamingService ?? throw new ArgumentNullException ("accountingSetNamingService");
			_accountingTemplateResolver = accountingTemplateResolver ?? throw new ArgumentNullException ("accountingTemplateResolver");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_transientPaneSuppressionService = transientPaneSuppressionService ?? throw new ArgumentNullException ("transientPaneSuppressionService");
			_accountingSetPresentationWaitService = accountingSetPresentationWaitService ?? throw new ArgumentNullException ("accountingSetPresentationWaitService");
			_accountingSetReadyShowBridge = accountingSetReadyShowBridge ?? throw new ArgumentNullException ("accountingSetReadyShowBridge");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (Workbook caseWorkbook)
		{
			if (caseWorkbook == null) {
				throw new ArgumentNullException ("caseWorkbook");
			}
			CaseContext caseContext = _caseContextFactory.CreateForDocumentCreate (caseWorkbook);
			if (caseContext == null || caseContext.CaseValues == null || caseContext.CaseValues.Count == 0) {
				throw new InvalidOperationException ("会計書類セット作成に必要な案件データを取得できませんでした。");
			}
			string text = ReadCustomerNameFromCase (caseContext);
			if (text.Length == 0) {
				throw new InvalidOperationException ("案件名を取得できませんでした。");
			}
			string valueText = BuildCustomerDisplayName (text, ReadCaseValue (caseContext, CustomerHonorificKey));
			string templatePath = _accountingTemplateResolver.ResolveTemplatePath (caseWorkbook);
			string outputFolderPath = _documentOutputService.ResolveWorkbookFolder (caseWorkbook);
			string outputPath = _accountingSetNamingService.BuildCaseOutputPath (caseWorkbook, outputFolderPath, text, templatePath);
			string workbookFullName = _excelInteropService.GetWorkbookFullName (caseWorkbook);
			_logger.Info ("Accounting set CASE create start. caseWorkbook=" + workbookFullName + ", customer=" + text + ", outputFolder=" + outputFolderPath + ", template=" + templatePath + ", output=" + outputPath);
			AccountingSetPresentationWaitService.WaitSession waitSession = null;
			Workbook workbook = null;
			try {
				waitSession = _accountingSetPresentationWaitService.ShowWaiting (Stopwatch.StartNew ());
				waitSession?.UpdateStage (AccountingSetPresentationWaitService.CreatingStageTitle);
				File.Copy (templatePath, outputPath, overwrite: false);
				_logger.Debug ("AccountingSetCreateService", "Template copied to output path.");
				_transientPaneSuppressionService.SuppressPath (outputPath, "AccountingSetCreateService.Execute");
				workbook = _accountingWorkbookService.OpenInCurrentApplication (outputPath);
				waitSession?.UpdateStage (AccountingSetPresentationWaitService.OpeningWorkbookStageTitle);
				_accountingWorkbookService.SetWorkbookWindowsVisible (workbook, visible: true);
				AccountingLawyerMappingResult accountingLawyerMappingResult;
				using (_accountingWorkbookService.BeginInitializationScope ()) {
					_excelInteropService.SetDocumentProperty (workbook, "CASEINFO_WORKBOOK_KIND", "ACCOUNTING_SET");
					_excelInteropService.SetDocumentProperty (workbook, "SOURCE_CASE_PATH", _excelInteropService.GetWorkbookFullName (caseWorkbook));
					_excelInteropService.SetDocumentProperty (workbook, "SYSTEM_ROOT", _excelInteropService.TryGetDocumentProperty (caseWorkbook, "SYSTEM_ROOT"));
					_excelInteropService.SetDocumentProperty (workbook, "NAME_RULE_A", _excelInteropService.TryGetDocumentProperty (caseWorkbook, "NAME_RULE_A"));
					_excelInteropService.SetDocumentProperty (workbook, "NAME_RULE_B", _excelInteropService.TryGetDocumentProperty (caseWorkbook, "NAME_RULE_B"));
					_logger.Debug ("AccountingSetCreateService", "Document property set. propertyName=CASEINFO_WORKBOOK_KIND, propertyValue=ACCOUNTING_SET");
					_logger.Debug ("AccountingSetCreateService", "Document property set. propertyName=SOURCE_CASE_PATH, propertyValue=" + workbookFullName);
					_logger.Debug ("AccountingSetCreateService", "Document properties copied from case. SYSTEM_ROOT=" + (_excelInteropService.TryGetDocumentProperty (caseWorkbook, "SYSTEM_ROOT") ?? string.Empty) + ", NAME_RULE_A=" + (_excelInteropService.TryGetDocumentProperty (caseWorkbook, "NAME_RULE_A") ?? string.Empty) + ", NAME_RULE_B=" + (_excelInteropService.TryGetDocumentProperty (caseWorkbook, "NAME_RULE_B") ?? string.Empty));
					string[] sheetNames = new string[3] { AccountingSetSpec.EstimateSheetName, AccountingSetSpec.InvoiceSheetName, AccountingSetSpec.ReceiptSheetName };
					_accountingWorkbookService.WriteSameValueToSheets (workbook, sheetNames, "A3", valueText);
					_accountingWorkbookService.WriteCell (workbook, AccountingSetSpec.AccountingRequestSheetName, "A3", valueText);
					string text5 = ReadCaseValue (caseContext, LawyerKey);
					_logger.Info ("Accounting set CASE lawyer source read. sourceKey=" + LawyerKey + ", lines=" + ToSingleLine (text5));
					accountingLawyerMappingResult = _accountingWorkbookService.ReflectLawyers (workbook, text5);
				}
				waitSession?.UpdateStage (AccountingSetPresentationWaitService.ApplyingInitialDataStageTitle);
				_transientPaneSuppressionService.ReleaseWorkbook (workbook, "AccountingSetCreateService.BeforeActivateInvoiceEntry");
				_logger.Info ("Accounting set CASE create pane handoff start. workbook=" + _excelInteropService.GetWorkbookFullName (workbook));
				_accountingWorkbookService.ActivateInvoiceEntry (workbook);
				_logger.Info ("Accounting set CASE create pane handoff activated. workbook=" + _excelInteropService.GetWorkbookFullName (workbook));
				_logger.Info ("Accounting set CASE create pane handoff before wait-ready. workbook=" + _excelInteropService.GetWorkbookFullName (workbook));
				waitSession?.UpdateStage (AccountingSetPresentationWaitService.ShowingInputScreenStageTitle);
				_accountingSetReadyShowBridge.ShowWorkbookTaskPaneWhenReady (workbook, "AccountingSetCreateService.Execute");
				waitSession?.Close ();
				_logger.Info ("Accounting set CASE create pane handoff queued. workbook=" + _excelInteropService.GetWorkbookFullName (workbook));
				if (accountingLawyerMappingResult.OverflowCount > 0) {
					MessageBox.Show ("入力できなかった代理人がいます。", "会計書類セット", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
				_logger.Info ("Accounting set CASE create completed. output=" + outputPath + ", missingMatch=" + accountingLawyerMappingResult.MissingMatchCount + ", overflow=" + accountingLawyerMappingResult.OverflowCount);
			} catch (Exception exception) {
				_logger.Error ("Accounting set CASE create failed. caseWorkbook=" + workbookFullName + ", template=" + templatePath + ", output=" + outputPath, exception);
				if (workbook != null) {
					try {
						workbook.Close (false, Type.Missing, Type.Missing);
					} catch (Exception ex) {
						_logger.Warn ("Accounting set CASE create cleanup close failed: " + ex.Message);
					}
				}
				_transientPaneSuppressionService.ReleasePath (outputPath, "AccountingSetCreateService.CleanupAfterFailure");
				try {
					if (_pathCompatibilityService.FileExistsSafe (outputPath)) {
						File.Delete (outputPath);
						_logger.Warn ("Accounting set CASE create cleanup deleted output: " + outputPath);
					}
				} catch (Exception ex2) {
					_logger.Warn ("Accounting set CASE create cleanup delete failed: " + ex2.Message);
				}
				throw;
			} finally {
				waitSession?.Dispose ();
			}
		}

		private static string ReadCaseValue (CaseContext caseContext, string key)
		{
			if (caseContext == null || caseContext.CaseValues == null || string.IsNullOrWhiteSpace (key)) {
				return string.Empty;
			}
			if (!caseContext.CaseValues.TryGetValue (key, out var value)) {
				return string.Empty;
			}
			return (value ?? string.Empty).Trim ();
		}

		private static string ReadCustomerNameFromCase (CaseContext caseContext)
		{
			return ReadCaseValue (caseContext, CustomerNameKey);
		}

		private static string BuildCustomerDisplayName (string customerName, string honorificText)
		{
			string text = (customerName ?? string.Empty).Trim ();
			string text2 = (honorificText ?? string.Empty).Trim ();
			if (text.Length == 0) {
				return string.Empty;
			}
			if (text2.Length == 0) {
				return text;
			}
			return text + " " + text2;
		}

		private static string ToSingleLine (string value)
		{
			return (value ?? string.Empty).Replace ("\r\n", " / ").Replace ('\r', ' ').Replace ('\n', ' ')
				.Trim ();
		}
	}
}
