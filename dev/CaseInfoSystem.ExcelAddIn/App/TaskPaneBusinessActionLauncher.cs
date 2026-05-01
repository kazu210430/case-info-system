using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneBusinessActionLauncher
    {
        private readonly DocumentCommandService _documentCommandService;
        private readonly DocumentNamePromptService _documentNamePromptService;

        internal TaskPaneBusinessActionLauncher(
            DocumentCommandService documentCommandService,
            DocumentNamePromptService documentNamePromptService)
        {
            _documentCommandService = documentCommandService ?? throw new ArgumentNullException(nameof(documentCommandService));
            _documentNamePromptService = documentNamePromptService ?? throw new ArgumentNullException(nameof(documentNamePromptService));
        }

        internal bool TryExecute(Excel.Workbook workbook, string actionKind, string key)
        {
            DocumentNameOverrideScope documentNameOverrideScope = null;
            try
            {
                if (string.Equals(actionKind, "doc", StringComparison.OrdinalIgnoreCase))
                {
                    bool shouldContinue = _documentNamePromptService.TryPrepare(workbook, key, out documentNameOverrideScope);
                    if (!shouldContinue)
                    {
                        return false;
                    }
                }

                _documentCommandService.Execute(workbook, actionKind, key);
                return true;
            }
            finally
            {
                if (documentNameOverrideScope != null)
                {
                    documentNameOverrideScope.Dispose();
                }
            }
        }
    }
}
