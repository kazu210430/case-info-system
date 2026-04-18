using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class WorkbookResetCommandService
	{
		private const string ProductTitle = "案件情報System";

		private readonly ExcelInteropService _excelInteropService;

		private readonly WorkbookRoleResolver _workbookRoleResolver;

		private readonly WorkbookResetDefinitionRepository _definitionRepository;

		private readonly KernelWorkbookLifecycleService _kernelWorkbookLifecycleService;

		private readonly Logger _logger;

		internal WorkbookResetCommandService (ExcelInteropService excelInteropService, WorkbookRoleResolver workbookRoleResolver, WorkbookResetDefinitionRepository definitionRepository, KernelWorkbookLifecycleService kernelWorkbookLifecycleService, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_definitionRepository = definitionRepository ?? throw new ArgumentNullException ("definitionRepository");
			_kernelWorkbookLifecycleService = kernelWorkbookLifecycleService ?? throw new ArgumentNullException ("kernelWorkbookLifecycleService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal WorkbookResetResult Execute (Workbook workbook)
		{
			if (workbook == null) {
				return new WorkbookResetResult {
					Success = false,
					Message = "対象ブックを取得できませんでした。"
				};
			}
			WorkbookResetDefinition workbookResetDefinition = ResolveDefinition (workbook);
			if (workbookResetDefinition == null) {
				return new WorkbookResetResult {
					Success = false,
					Message = "Kernel または Base ブックで実行してください。"
				};
			}
			string workbookName = _excelInteropService.GetWorkbookName (workbook);
			DialogResult dialogResult = MessageBox.Show (BuildConfirmationMessage (workbookResetDefinition, workbookName), "案件情報System", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
			if (dialogResult != DialogResult.OK) {
				return new WorkbookResetResult {
					Success = false,
					Message = "配布前リセットをキャンセルしました。"
				};
			}
			try {
				ApplyFixedValues (workbook, workbookResetDefinition);
				ClearDynamicProperties (workbook, workbookResetDefinition);
				SaveWorkbook (workbook, workbookResetDefinition);
				string message = workbookResetDefinition.TargetName + " の配布前リセットが完了しました。";
				_logger.Info ("Workbook reset completed. target=" + workbookResetDefinition.TargetName + ", workbook=" + _excelInteropService.GetWorkbookFullName (workbook));
				return new WorkbookResetResult {
					Success = true,
					Message = message
				};
			} catch (Exception exception) {
				_logger.Error ("Workbook reset failed.", exception);
				return new WorkbookResetResult {
					Success = false,
					Message = "配布前リセットに失敗しました。ログを確認してください。"
				};
			}
		}

		internal void ShowResult (WorkbookResetResult result)
		{
			if (result != null && !string.IsNullOrWhiteSpace (result.Message)) {
				MessageBoxIcon icon = (result.Success ? MessageBoxIcon.Asterisk : MessageBoxIcon.Exclamation);
				MessageBox.Show (result.Message, "案件情報System", MessageBoxButtons.OK, icon);
			}
		}

		private WorkbookResetDefinition ResolveDefinition (Workbook workbook)
		{
			if (_workbookRoleResolver.IsKernelWorkbook (workbook)) {
				return _definitionRepository.GetKernelDefinition ();
			}
			if (_workbookRoleResolver.IsBaseWorkbook (workbook)) {
				return _definitionRepository.GetBaseDefinition ();
			}
			return null;
		}

		private void ApplyFixedValues (Workbook workbook, WorkbookResetDefinition definition)
		{
			foreach (KeyValuePair<string, string> fixedValue in definition.FixedValues) {
				_excelInteropService.SetDocumentProperty (workbook, fixedValue.Key, fixedValue.Value);
			}
		}

		private void ClearDynamicProperties (Workbook workbook, WorkbookResetDefinition definition)
		{
			if (definition.ClearPrefixNames == null || definition.ClearPrefixNames.Count == 0) {
				return;
			}
			IReadOnlyList<KeyValuePair<string, string>> customDocumentProperties = _excelInteropService.GetCustomDocumentProperties (workbook);
			foreach (KeyValuePair<string, string> item in customDocumentProperties) {
				if (ShouldClearProperty (definition.ClearPrefixNames, item.Key)) {
					_excelInteropService.SetDocumentProperty (workbook, item.Key, string.Empty);
				}
			}
		}

		private void SaveWorkbook (Workbook workbook, WorkbookResetDefinition definition)
		{
			if (workbook == null) {
				throw new ArgumentNullException ("workbook");
			}
			if (definition != null && string.Equals (definition.TargetName, "Kernel", StringComparison.OrdinalIgnoreCase)) {
				using (_kernelWorkbookLifecycleService.SuppressBeforeSaveDocPropSynchronization ("WorkbookResetCommandService.Execute")) {
					workbook.Save ();
					return;
				}
			}
			workbook.Save ();
		}

		private static string BuildConfirmationMessage (WorkbookResetDefinition definition, string workbookName)
		{
			return (workbookName ?? string.Empty) + " に対して配布前リセットを実行します。" + Environment.NewLine + "この処理は DocProp を初期値へ戻し、最後に保存します。" + Environment.NewLine + "続行しますか？";
		}

		private static bool ShouldClearProperty (IReadOnlyList<string> clearPrefixNames, string propertyName)
		{
			if (clearPrefixNames == null || string.IsNullOrWhiteSpace (propertyName)) {
				return false;
			}
			for (int i = 0; i < clearPrefixNames.Count; i++) {
				string value = clearPrefixNames [i];
				if (!string.IsNullOrWhiteSpace (value) && propertyName.StartsWith (value, StringComparison.OrdinalIgnoreCase)) {
					return true;
				}
			}
			return false;
		}
	}
}
