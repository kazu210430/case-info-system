using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    // ThisAddIn の lifecycle 呼び出し位置を変えずに、Excel application event の配線だけを担当する。
    internal sealed class ApplicationEventSubscriptionService
    {
        private readonly Excel.AppEvents_Event _applicationEvents;
        private readonly Excel.AppEvents_WorkbookOpenEventHandler _workbookOpenHandler;
        private readonly Excel.AppEvents_WorkbookActivateEventHandler _workbookActivateHandler;
        private readonly Excel.AppEvents_WorkbookBeforeSaveEventHandler _workbookBeforeSaveHandler;
        private readonly Excel.AppEvents_WorkbookBeforeCloseEventHandler _workbookBeforeCloseHandler;
        private readonly Excel.AppEvents_WindowActivateEventHandler _windowActivateHandler;
        private readonly Excel.AppEvents_SheetActivateEventHandler _sheetActivateHandler;
        private readonly Excel.AppEvents_SheetSelectionChangeEventHandler _sheetSelectionChangeHandler;
        private readonly Excel.AppEvents_SheetChangeEventHandler _sheetChangeHandler;
        private readonly Excel.AppEvents_AfterCalculateEventHandler _afterCalculateHandler;
        private readonly bool _subscribeSheetActivate;
        private readonly bool _subscribeSheetSelectionChange;
        private readonly bool _subscribeSheetChange;
        private bool _isSubscribed;

        internal ApplicationEventSubscriptionService(
            Excel.Application application,
            Excel.AppEvents_WorkbookOpenEventHandler workbookOpenHandler,
            Excel.AppEvents_WorkbookActivateEventHandler workbookActivateHandler,
            Excel.AppEvents_WorkbookBeforeSaveEventHandler workbookBeforeSaveHandler,
            Excel.AppEvents_WorkbookBeforeCloseEventHandler workbookBeforeCloseHandler,
            Excel.AppEvents_WindowActivateEventHandler windowActivateHandler,
            Excel.AppEvents_SheetActivateEventHandler sheetActivateHandler,
            Excel.AppEvents_SheetSelectionChangeEventHandler sheetSelectionChangeHandler,
            Excel.AppEvents_SheetChangeEventHandler sheetChangeHandler,
            Excel.AppEvents_AfterCalculateEventHandler afterCalculateHandler,
            bool subscribeSheetActivate,
            bool subscribeSheetSelectionChange,
            bool subscribeSheetChange)
        {
            if (application == null)
            {
                throw new ArgumentNullException(nameof(application));
            }

            _applicationEvents = application as Excel.AppEvents_Event
                ?? throw new ArgumentException("Excel application event interface could not be resolved.", nameof(application));
            _workbookOpenHandler = workbookOpenHandler ?? throw new ArgumentNullException(nameof(workbookOpenHandler));
            _workbookActivateHandler = workbookActivateHandler ?? throw new ArgumentNullException(nameof(workbookActivateHandler));
            _workbookBeforeSaveHandler = workbookBeforeSaveHandler ?? throw new ArgumentNullException(nameof(workbookBeforeSaveHandler));
            _workbookBeforeCloseHandler = workbookBeforeCloseHandler ?? throw new ArgumentNullException(nameof(workbookBeforeCloseHandler));
            _windowActivateHandler = windowActivateHandler ?? throw new ArgumentNullException(nameof(windowActivateHandler));
            _sheetActivateHandler = sheetActivateHandler ?? throw new ArgumentNullException(nameof(sheetActivateHandler));
            _sheetSelectionChangeHandler = sheetSelectionChangeHandler ?? throw new ArgumentNullException(nameof(sheetSelectionChangeHandler));
            _sheetChangeHandler = sheetChangeHandler ?? throw new ArgumentNullException(nameof(sheetChangeHandler));
            _afterCalculateHandler = afterCalculateHandler ?? throw new ArgumentNullException(nameof(afterCalculateHandler));
            _subscribeSheetActivate = subscribeSheetActivate;
            _subscribeSheetSelectionChange = subscribeSheetSelectionChange;
            _subscribeSheetChange = subscribeSheetChange;
        }

        internal void Subscribe()
        {
            if (_isSubscribed)
            {
                return;
            }

            _applicationEvents.WorkbookOpen += _workbookOpenHandler;
            _applicationEvents.WorkbookActivate += _workbookActivateHandler;
            _applicationEvents.WorkbookBeforeSave += _workbookBeforeSaveHandler;
            _applicationEvents.WorkbookBeforeClose += _workbookBeforeCloseHandler;
            _applicationEvents.WindowActivate += _windowActivateHandler;
            if (_subscribeSheetActivate)
            {
                _applicationEvents.SheetActivate += _sheetActivateHandler;
            }

            if (_subscribeSheetSelectionChange)
            {
                _applicationEvents.SheetSelectionChange += _sheetSelectionChangeHandler;
            }

            if (_subscribeSheetChange)
            {
                _applicationEvents.SheetChange += _sheetChangeHandler;
            }

            _applicationEvents.AfterCalculate += _afterCalculateHandler;
            _isSubscribed = true;
        }

        internal void Unsubscribe()
        {
            if (!_isSubscribed)
            {
                return;
            }

            _applicationEvents.WorkbookOpen -= _workbookOpenHandler;
            _applicationEvents.WorkbookActivate -= _workbookActivateHandler;
            _applicationEvents.WorkbookBeforeSave -= _workbookBeforeSaveHandler;
            _applicationEvents.WorkbookBeforeClose -= _workbookBeforeCloseHandler;
            _applicationEvents.WindowActivate -= _windowActivateHandler;
            if (_subscribeSheetActivate)
            {
                _applicationEvents.SheetActivate -= _sheetActivateHandler;
            }

            if (_subscribeSheetSelectionChange)
            {
                _applicationEvents.SheetSelectionChange -= _sheetSelectionChangeHandler;
            }

            if (_subscribeSheetChange)
            {
                _applicationEvents.SheetChange -= _sheetChangeHandler;
            }

            _applicationEvents.AfterCalculate -= _afterCalculateHandler;
            _isSubscribed = false;
        }
    }
}
