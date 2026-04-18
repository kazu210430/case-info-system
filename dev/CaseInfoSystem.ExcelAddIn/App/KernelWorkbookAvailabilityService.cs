using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookAvailabilityService
    {
        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly ExcelInteropService _excelInteropService;
        private readonly KernelHomeCoordinator _kernelHomeCoordinator;
        private readonly Action<string> _showKernelHomePlaceholderWithExternalWorkbookSuppression;
        private readonly Logger _logger;

        internal KernelWorkbookAvailabilityService(
            KernelWorkbookService kernelWorkbookService,
            ExcelInteropService excelInteropService,
            KernelHomeCoordinator kernelHomeCoordinator,
            Action<string> showKernelHomePlaceholderWithExternalWorkbookSuppression,
            Logger logger)
        {
            _kernelWorkbookService = kernelWorkbookService;
            _excelInteropService = excelInteropService;
            _kernelHomeCoordinator = kernelHomeCoordinator;
            _showKernelHomePlaceholderWithExternalWorkbookSuppression = showKernelHomePlaceholderWithExternalWorkbookSuppression;
            _logger = logger;
        }

        internal void Handle(string eventName, Excel.Workbook workbook, KernelHomeForm kernelHomeForm)
        {
            try
            {
                if (_kernelWorkbookService == null)
                {
                    return;
                }

                if (workbook != null && !_kernelWorkbookService.IsKernelWorkbook(workbook))
                {
                    _logger.Info("HandleKernelWorkbookBecameAvailable skipped for non-kernel workbook. eventName=" + (eventName ?? string.Empty) + ", workbook=" + _excelInteropService.GetWorkbookFullName(workbook));
                    return;
                }

                // 処理ブロック: suppression 判定の consume ポイント。
                // 判定と消費は coordinator の責務で、このメソッドは結果に応じた後続動作だけを調停する。
                _logger?.Info("[Suppression:Check] event=" + eventName);
                if (_kernelHomeCoordinator.ShouldSuppressKernelHomeDisplay(eventName))
                {
                    _logger?.Info("[Suppression:Consumed] event=" + eventName);
                    _logger.Info((eventName ?? string.Empty) + " suppressed kernel home display.");
                    return;
                }

                // 処理ブロック: HOME 表示抑止が成立しなかった場合のみ、イベント種別と現在状態に応じて表示分岐する。
                if (kernelHomeForm == null || kernelHomeForm.IsDisposed)
                {
                    if (_kernelHomeCoordinator.ShouldAutoShowKernelHomeForEvent(eventName, workbook))
                    {
                        _logger.Info("Kernel HOME requested from " + (eventName ?? string.Empty));
                        _showKernelHomePlaceholderWithExternalWorkbookSuppression("HandleKernelWorkbookBecameAvailable." + (eventName ?? string.Empty));
                    }

                    return;
                }

                // 処理ブロック: 既存 HOME 表示中の UI 更新。
                if (kernelHomeForm.Visible && _kernelHomeCoordinator.ShouldReloadVisibleKernelHomeForEvent(eventName, workbook))
                {
                    kernelHomeForm.ReloadSettings();
                    _logger.Info("KernelHomeForm reloaded after " + (eventName ?? string.Empty));
                }
            }
            catch (Exception ex)
            {
                _logger.Error("HandleKernelWorkbookBecameAvailable failed. eventName=" + (eventName ?? string.Empty), ex);
            }
        }
    }
}
