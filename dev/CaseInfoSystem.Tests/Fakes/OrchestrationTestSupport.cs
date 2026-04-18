using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests.Fakes
{
    internal static class OrchestrationTestSupport
    {
        internal static Logger CreateLogger(List<string> messages)
        {
            return new Logger(message => messages.Add(message));
        }

        internal static KernelCaseInteractionState CreateKernelCaseInteractionState(List<string> messages)
        {
            return new KernelCaseInteractionState(CreateLogger(messages));
        }

        internal static TaskPaneHost CreateTaskPaneHost(System.Windows.Forms.UserControl control, string windowKey)
        {
            return new TaskPaneHost(new CaseInfoSystem.ExcelAddIn.ThisAddIn(), new Excel.Window { Hwnd = windowKey.GetHashCode() }, control, (ITaskPaneView)control, windowKey);
        }
    }
}
