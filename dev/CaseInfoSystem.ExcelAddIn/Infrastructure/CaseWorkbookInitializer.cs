using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class CaseWorkbookInitializer
	{
		private const string HomeSheetCodeName = "shHOME";

		private const string HomeSheetName = "ホーム";

		private const string TemplateFolderName = "雛形";

		private const string CustomerNameFieldKey = "顧客_名前";

		private readonly ExcelInteropService _excelInteropService;

		private readonly CaseTemplateSnapshotService _caseTemplateSnapshotService;

		private readonly CaseListFieldDefinitionRepository _caseListFieldDefinitionRepository;

		internal CaseWorkbookInitializer (ExcelInteropService excelInteropService, CaseTemplateSnapshotService caseTemplateSnapshotService, CaseListFieldDefinitionRepository caseListFieldDefinitionRepository)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_caseTemplateSnapshotService = caseTemplateSnapshotService ?? throw new ArgumentNullException ("caseTemplateSnapshotService");
			_caseListFieldDefinitionRepository = caseListFieldDefinitionRepository ?? throw new ArgumentNullException ("caseListFieldDefinitionRepository");
		}

		internal void InitializeForVisibleCreate (Workbook kernelWorkbook, Workbook caseWorkbook, KernelCaseCreationPlan plan)
		{
			InitializeCore (kernelWorkbook, caseWorkbook, plan, "1", "0");
		}

		internal void InitializeForHiddenCreate (Workbook kernelWorkbook, Workbook caseWorkbook, KernelCaseCreationPlan plan)
		{
			InitializeCore (kernelWorkbook, caseWorkbook, plan, "0", "1");
		}

		internal void CompleteVisibleCreateStartupState (Workbook caseWorkbook)
		{
			_excelInteropService.SetDocumentProperty (caseWorkbook, "KERNEL_JUST_CREATED", "0");
			_excelInteropService.SetDocumentProperty (caseWorkbook, "TASKPANE_READY", "1");
		}

		private void InitializeCore (Workbook kernelWorkbook, Workbook caseWorkbook, KernelCaseCreationPlan plan, string kernelJustCreated, string taskPaneReady)
		{
			SetCoreDocumentProperties (caseWorkbook, plan);
			_excelInteropService.SetDocumentProperty (caseWorkbook, "KERNEL_JUST_CREATED", kernelJustCreated);
			_excelInteropService.SetDocumentProperty (caseWorkbook, "TASKPANE_READY", taskPaneReady);
			ApplyCustomerToHomeSheet (kernelWorkbook, caseWorkbook, plan.CustomerName);
			_caseTemplateSnapshotService.SyncMasterVersionFromKernel (kernelWorkbook, caseWorkbook);
			_caseTemplateSnapshotService.PromoteEmbeddedSnapshotToCaseCache (caseWorkbook);
		}

		private void SetCoreDocumentProperties (Workbook caseWorkbook, KernelCaseCreationPlan plan)
		{
			_excelInteropService.SetDocumentProperty (caseWorkbook, "ROLE", "CASE");
			_excelInteropService.SetDocumentProperty (caseWorkbook, "SYSTEM_ROOT", plan.SystemRoot ?? string.Empty);
			_excelInteropService.SetDocumentProperty (caseWorkbook, "WORD_TEMPLATE_DIR", (plan.SystemRoot ?? string.Empty) + "\\雛形");
			_excelInteropService.SetDocumentProperty (caseWorkbook, "NAME_RULE_A", plan.NameRuleA ?? string.Empty);
			_excelInteropService.SetDocumentProperty (caseWorkbook, "NAME_RULE_B", plan.NameRuleB ?? string.Empty);
		}

		private void ApplyCustomerToHomeSheet (Workbook kernelWorkbook, Workbook caseWorkbook, string customerName)
		{
			if (string.IsNullOrWhiteSpace (customerName)) {
				return;
			}
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (caseWorkbook, "shHOME");
			if (worksheet == null) {
				try {
					worksheet = caseWorkbook.Worksheets ["ホーム"] as Worksheet;
				} catch {
					worksheet = caseWorkbook.Worksheets [1] as Worksheet;
				}
			}
			if (worksheet == null) {
				return;
			}
			IReadOnlyDictionary<string, CaseListFieldDefinition> readOnlyDictionary = _caseListFieldDefinitionRepository.LoadDefinitions (kernelWorkbook);
			readOnlyDictionary.TryGetValue ("顧客_名前", out var value);
			if (!_excelInteropService.TryWriteFieldValue (caseWorkbook, worksheet, value, customerName)) {
				return;
			}
			try {
				if (caseWorkbook.Windows.Count > 0) {
					caseWorkbook.Windows [1].ScrollColumn = 1;
				}
			} catch {
			}
		}
	}
}
