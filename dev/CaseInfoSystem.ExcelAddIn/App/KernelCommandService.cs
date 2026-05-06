using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelCommandService
	{
		internal const string SheetCommandOpenHome = "open-home";

		internal const string SheetCommandReflectAccountingSet = "reflect-accounting-set";

		internal const string SheetCommandReflectBaseHome = "reflect-base-home";

		private readonly KernelWorkbookService _kernelWorkbookService;

		private readonly KernelUserDataReflectionService _kernelUserDataReflectionService;

		private readonly KernelTemplateSyncService _kernelTemplateSyncService;

		private readonly Action _showKernelHomeAction;

		private readonly Logger _logger;

		internal KernelCommandService (KernelWorkbookService kernelWorkbookService, KernelUserDataReflectionService kernelUserDataReflectionService, KernelTemplateSyncService kernelTemplateSyncService, Action showKernelHomeAction, Logger logger)
		{
			_kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException ("kernelWorkbookService");
			_kernelUserDataReflectionService = kernelUserDataReflectionService ?? throw new ArgumentNullException ("kernelUserDataReflectionService");
			_kernelTemplateSyncService = kernelTemplateSyncService ?? throw new ArgumentNullException ("kernelTemplateSyncService");
			_showKernelHomeAction = showKernelHomeAction ?? throw new ArgumentNullException ("showKernelHomeAction");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void Execute (WorkbookContext context, string actionId)
		{
			if (!string.IsNullOrWhiteSpace (actionId)) {
				_logger.Info ("Kernel pane action requested. actionId=" + actionId);
				switch (actionId) {
				case "open-home":
					_kernelWorkbookService.BindHomeWorkbook (context);
					_showKernelHomeAction ();
					break;
				case "open-user-info":
					ExecuteSheetNavigation (context, "shUserData", "ユーザー情報");
					break;
				case "open-template-list":
					ExecuteSheetNavigation (context, "shMasterList", "雛形一覧");
					break;
				case "open-case-list":
					ExecuteSheetNavigation (context, "shCaseList", "案件一覧");
					break;
				case "register-user-info":
					ExecuteRegisterUserInfo (context);
					break;
				case "reflect-template":
					ExecuteReflectTemplate (context);
					break;
				default:
					UserErrorService.ShowOkNotification ("未対応の操作です。actionId=" + actionId, "案件情報System", MessageBoxIcon.Asterisk);
					break;
				}
			}
		}

		internal void ExecuteSheetCommand (string commandId)
		{
			ExecuteSheetCommand (null, commandId);
		}

		internal void ExecuteSheetCommand (WorkbookContext context, string commandId)
		{
			if (!string.IsNullOrWhiteSpace (commandId)) {
				_logger.Info ("Kernel sheet command requested. commandId=" + commandId);
				switch (commandId.Trim ()) {
				case "open-home":
					_kernelWorkbookService.ClearHomeWorkbookBinding ("KernelCommandService.ExecuteSheetCommand.OpenHome");
					_showKernelHomeAction ();
					break;
				case "reflect-accounting-set":
					ExecuteReflectAccountingSetOnly (context);
					break;
				case "reflect-base-home":
					ExecuteReflectBaseHomeOnly (context);
					break;
				default:
					_logger.Warn ("Kernel sheet command ignored. commandId=" + commandId);
					break;
				}
			}
		}

		private void ExecuteSheetNavigation (WorkbookContext context, string codeName, string featureName)
		{
			if (!_kernelWorkbookService.TryShowSheetByCodeName (context, codeName, "KernelTaskPane." + (codeName ?? string.Empty))) {
				_logger.Warn ("Kernel sheet navigation failed. feature=" + featureName + ", codeName=" + (codeName ?? string.Empty));
				UserErrorService.ShowOkNotification (featureName + " を開けませんでした。ログを確認してください。", "案件情報System", MessageBoxIcon.Exclamation);
			}
		}

		private void ExecuteRegisterUserInfo (WorkbookContext context)
		{
			try {
				_kernelUserDataReflectionService.ReflectAll (context);
				UserErrorService.ShowOkNotification ("ユーザー情報を反映しました", "案件情報System", MessageBoxIcon.Asterisk);
			} catch (Exception exception) {
				_logger.Error ("ExecuteRegisterUserInfo failed.", exception);
				UserErrorService.ShowOkNotification ("ユーザー情報登録を実行できませんでした。ログを確認してください。", "案件情報System", MessageBoxIcon.Exclamation);
			}
		}

		private void ExecuteReflectTemplate (WorkbookContext context)
		{
			try {
				KernelTemplateSyncResult kernelTemplateSyncResult = _kernelTemplateSyncService.Execute (context);
				TemplateRegistrationResultForm.ShowNotice ("案件情報System", kernelTemplateSyncResult.Message);
			} catch (Exception exception) {
				_logger.Error ("ExecuteReflectTemplate failed.", exception);
				UserErrorService.ShowOkNotification ("雛形登録・更新を実行できませんでした。ログを確認してください。", "案件情報System", MessageBoxIcon.Exclamation);
			}
		}

		private void ExecuteReflectAccountingSetOnly (WorkbookContext context)
		{
			try {
				_kernelUserDataReflectionService.ReflectToAccountingSetOnly (context);
			} catch (Exception exception) {
				_logger.Error ("ExecuteReflectAccountingSetOnly failed.", exception);
				UserErrorService.ShowOkNotification ("会計書類セットへの転記でエラーが発生しました。ログを確認してください。", "案件情報System", MessageBoxIcon.Exclamation);
			}
		}

		private void ExecuteReflectBaseHomeOnly (WorkbookContext context)
		{
			try {
				_kernelUserDataReflectionService.ReflectToBaseHomeOnly (context);
			} catch (Exception exception) {
				_logger.Error ("ExecuteReflectBaseHomeOnly failed.", exception);
				UserErrorService.ShowOkNotification ("Baseホームへの転記でエラーが発生しました。ログを確認してください。", "案件情報System", MessageBoxIcon.Exclamation);
			}
		}
	}
}
