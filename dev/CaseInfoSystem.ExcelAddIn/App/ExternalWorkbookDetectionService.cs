using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class ExternalWorkbookDetectionService
    {
        private readonly WorkbookRoleResolver _workbookRoleResolver;
        private readonly KernelCaseInteractionState _kernelCaseInteractionState;
        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly TransientPaneSuppressionService _transientPaneSuppressionService;
        private readonly IExcelInteropService _excelInteropService;
        private readonly Logger _logger;

        internal ExternalWorkbookDetectionService(
            WorkbookRoleResolver workbookRoleResolver,
            KernelCaseInteractionState kernelCaseInteractionState,
            KernelWorkbookService kernelWorkbookService,
            TransientPaneSuppressionService transientPaneSuppressionService,
            IExcelInteropService excelInteropService,
            Logger logger)
        {
            _workbookRoleResolver = workbookRoleResolver;
            _kernelCaseInteractionState = kernelCaseInteractionState;
            _kernelWorkbookService = kernelWorkbookService;
            _transientPaneSuppressionService = transientPaneSuppressionService;
            _excelInteropService = excelInteropService;
            _logger = logger;
        }

        internal void Handle(
            Excel.Workbook workbook,
            string eventName,
            KernelHomeForm kernelHomeForm,
            Func<string, bool, bool> isKernelHomeSuppressionActive,
            ref bool kernelHomeExternalCloseBusy,
            ref bool kernelHomeExternalCloseRequested)
        {
            try
            {
                bool isCaseWorkbook = _workbookRoleResolver != null && _workbookRoleResolver.IsCaseWorkbook(workbook);
                bool isKernelCaseCreationFlowActive = _kernelCaseInteractionState != null
                    && _kernelCaseInteractionState.IsKernelCaseCreationFlowActive;
                if (!KernelHomeExternalClosePolicy.ShouldCloseKernelHome(isKernelCaseCreationFlowActive))
                {
                    _logger?.Info(
                        "Kernel HOME external workbook detection ignored to preserve cross-workbook state. eventName="
                        + (eventName ?? string.Empty)
                        + ", workbook="
                        + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));
                    return;
                }

                if (kernelHomeExternalCloseBusy || kernelHomeExternalCloseRequested)
                {
                    return;
                }

                if (_kernelWorkbookService == null || kernelHomeForm == null || kernelHomeForm.IsDisposed || !kernelHomeForm.Visible)
                {
                    return;
                }

                if (workbook == null || _kernelWorkbookService.IsKernelWorkbook(workbook))
                {
                    return;
                }

                // 処理ブロック: ここは抑止状態の参照のみを行い、外部 workbook 検知を無視するか判定する。抑止は消費しない。
                _logger?.Info("[Suppression:CheckOnly] event=" + eventName);
                if (isKernelHomeSuppressionActive != null && isKernelHomeSuppressionActive(eventName, false))
                {
                    _logger?.Info(
                        "Kernel HOME external workbook detected but ignored during kernel-home suppression. eventName="
                        + (eventName ?? string.Empty)
                        + ", workbook="
                        + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));
                    return;
                }

                if (_transientPaneSuppressionService != null && _transientPaneSuppressionService.IsSuppressed(workbook))
                {
                    _logger?.Info(
                        "Kernel HOME external workbook detected but ignored during transient suppression. eventName="
                        + (eventName ?? string.Empty)
                        + ", workbook="
                        + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));
                    return;
                }

                kernelHomeExternalCloseBusy = true;
                kernelHomeExternalCloseRequested = true;
                if (isCaseWorkbook)
                {
                    // 処理ブロック: 既存 CASE オープン起点では、Kernel の内部更新内容を自動保存して保存確認を抑止する。
                    kernelHomeForm.PrepareForExistingCaseOpenClose();
                }

                _logger?.Info(
                    "Kernel HOME external workbook detected. eventName="
                    + (eventName ?? string.Empty)
                    + ", workbook="
                    + (_excelInteropService == null ? string.Empty : _excelInteropService.GetWorkbookFullName(workbook)));

                kernelHomeForm.Close();
            }
            catch (Exception ex)
            {
                kernelHomeExternalCloseRequested = false;
                _logger?.Error("HandleExternalWorkbookDetected failed. eventName=" + (eventName ?? string.Empty), ex);
            }
            finally
            {
                kernelHomeExternalCloseBusy = false;
            }
        }
    }
}
