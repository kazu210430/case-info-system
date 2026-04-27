using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Microsoft.Office.Tools
{
    public class CustomTaskPane
    {
        public int Width { get; set; }

        public bool Visible { get; set; }
    }
}

namespace Microsoft.WindowsAPICodePack.Dialogs
{
    public enum CommonFileDialogResult
    {
        Ok,
        Cancel
    }

    public sealed class CommonOpenFileDialog : IDisposable
    {
        public bool IsFolderPicker { get; set; }

        public bool Multiselect { get; set; }

        public string Title { get; set; }

        public bool EnsurePathExists { get; set; }

        public bool AllowNonFileSystemItems { get; set; }

        public string InitialDirectory { get; set; }

        public string DefaultDirectory { get; set; }

        public string FileName { get; set; }

        public CommonFileDialogResult ShowDialog()
        {
            return CommonFileDialogResult.Cancel;
        }

        public void Dispose()
        {
        }
    }
}

namespace Microsoft.Office.Interop.Excel
{
    public enum XlWindowState
    {
        xlNormal = 0,
        xlMinimized = -4140,
        xlMaximized = -4137
    }

    public enum XlFileFormat
    {
        xlOpenXMLWorkbookMacroEnabled = 52
    }

    public class Application
    {
        public bool DisplayAlerts { get; set; }

        public bool EnableEvents { get; set; }

        public bool ScreenUpdating { get; set; } = true;

        public bool Visible { get; set; }

        public int Hwnd { get; set; }

        public Workbook ActiveWorkbook { get; set; }

        public Window ActiveWindow { get; set; }

        public Workbooks Workbooks { get; } = new Workbooks();

        public int QuitCallCount { get; private set; }

        public Action QuitBehavior { get; set; }

        public void Quit()
        {
            QuitCallCount++;
            QuitBehavior?.Invoke();
        }
    }

    public class Workbooks : List<Workbook>
    {
        public Workbook Open(string filename, object UpdateLinks = null, bool ReadOnly = false)
        {
            var workbook = new Workbook
            {
                FullName = filename ?? string.Empty,
                Name = Path.GetFileName(filename ?? string.Empty),
                Path = Path.GetDirectoryName(filename ?? string.Empty) ?? string.Empty
            };
            Add(workbook);
            return workbook;
        }

        public Workbook Open(
            string filename,
            object UpdateLinks,
            bool ReadOnly,
            object Format,
            object Password,
            object WriteResPassword,
            object IgnoreReadOnlyRecommended,
            object Origin,
            object Delimiter,
            object Editable,
            object Notify,
            object Converter,
            object AddToMru,
            object Local,
            object CorruptLoad)
        {
            return Open(filename, UpdateLinks, ReadOnly);
        }
    }

    public class Workbook
    {
        public Application Application { get; set; }

        public string FullName { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public string Path { get; set; } = string.Empty;

        public bool Saved { get; set; }

        public XlFileFormat FileFormat { get; set; }

        public WorkbookWindows Windows { get; } = new WorkbookWindows();

        public Worksheet ActiveSheet { get; set; }

        public Worksheets Worksheets { get; } = new Worksheets();

        public object CustomDocumentProperties { get; set; }

        public int SaveCallCount { get; private set; }

        public Action SaveBehavior { get; set; }

        public int CloseCallCount { get; private set; }

        public Action CloseBehavior { get; set; }

        public void Save()
        {
            SaveCallCount++;
            if (SaveBehavior != null)
            {
                SaveBehavior();
                return;
            }

            Saved = true;
        }

        public void Close(bool SaveChanges = false, object Filename = null, object RouteWorkbook = null)
        {
            CloseCallCount++;
            CloseBehavior?.Invoke();
        }

        public void Activate()
        {
            if (Application != null)
            {
                Application.ActiveWorkbook = this;
                Application.ActiveWindow = Windows.Count > 0 ? Windows[1] : null;
            }
        }

        public Window NewWindow()
        {
            var window = new Window();
            Windows.Add(window);
            if (Application != null)
            {
                Application.ActiveWorkbook = this;
                Application.ActiveWindow = window;
            }
            return window;
        }
    }

    public class WorkbookWindows : List<Window>
    {
        public new int Count => base.Count;

        public new Window this[int index]
        {
            get => base[index - 1];
            set => base[index - 1] = value;
        }
    }

    public class Window
    {
        public bool Visible { get; set; } = true;

        public int Hwnd { get; set; }

        public bool Activated { get; private set; }

        public bool FreezePanes { get; set; }

        public int SplitRow { get; set; }

        public int SplitColumn { get; set; }

        public int ScrollRow { get; set; }

        public int ScrollColumn { get; set; }

        public XlWindowState WindowState { get; set; }

        public void Activate()
        {
            Activated = true;
        }
    }

    public class Worksheet
    {
        public string CodeName { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public object Parent { get; set; }

        public WorksheetCellCollection Cells { get; } = new WorksheetCellCollection();

        public void Activate()
        {
        }
    }

    public class Worksheets : List<Worksheet>
    {
        public new Worksheet this[int index]
        {
            get => base[index - 1];
            set => base[index - 1] = value;
        }

        public Worksheet this[string name] => this.FirstOrDefault(worksheet => string.Equals(worksheet?.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    public class Range
    {
        public object Value2 { get; set; }

        public int Start { get; set; }

        public int End { get; set; }

        public string Text
        {
            get => Convert.ToString(Value2) ?? string.Empty;
            set => Value2 = value;
        }
    }

    public sealed class WorksheetCellCollection
    {
        private readonly Dictionary<string, Range> _cells = new Dictionary<string, Range>(StringComparer.OrdinalIgnoreCase);

        private readonly HashSet<string> _throwKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public Range this[object row, object column]
        {
            get
            {
                string key = BuildKey(row, column);
                if (_throwKeys.Contains(key))
                {
                    throw new InvalidOperationException("Configured cell access failure.");
                }

                if (!_cells.TryGetValue(key, out Range range))
                {
                    range = new Range();
                    _cells[key] = range;
                }

                return range;
            }
        }

        public void SetValue(object row, object column, object value)
        {
            this[row, column].Value2 = value;
        }

        public void ThrowOnAccess(object row, object column)
        {
            _throwKeys.Add(BuildKey(row, column));
        }

        private static string BuildKey(object row, object column)
        {
            return string.Concat(Convert.ToString(row) ?? string.Empty, "|", Convert.ToString(column) ?? string.Empty);
        }
    }
}

namespace CaseInfoSystem.ExcelAddIn
{
    internal static class Globals
    {
        internal static ThisAddIn ThisAddIn { get; set; } = new ThisAddIn();
    }

    internal sealed class ThisAddIn
    {
        internal Action<Action> RunWithScreenUpdatingSuspendedHandler { get; set; }

        internal Func<string, IDisposable> SuppressTaskPaneRefreshHandler { get; set; }

        internal Action<string> RefreshActiveTaskPaneHandler { get; set; }

        internal Action<CaseInfoSystem.ExcelAddIn.App.TaskPaneDisplayRequest, Microsoft.Office.Interop.Excel.Workbook, Microsoft.Office.Interop.Excel.Window> RequestTaskPaneDisplayForTargetWindowHandler { get; set; }

        internal Func<string, string, bool> ShowKernelSheetAndRefreshPaneHandler { get; set; }

        internal Action<Microsoft.Office.Interop.Excel.Workbook, string> ShowWorkbookTaskPaneWhenReadyHandler { get; set; }

        internal Microsoft.Office.Tools.CustomTaskPane CreateTaskPane(Microsoft.Office.Interop.Excel.Window window, UserControl control)
        {
            return new Microsoft.Office.Tools.CustomTaskPane();
        }

        internal void RemoveTaskPane(Microsoft.Office.Tools.CustomTaskPane pane)
        {
        }

        internal void RunWithScreenUpdatingSuspended(Action action)
        {
            if (RunWithScreenUpdatingSuspendedHandler != null)
            {
                RunWithScreenUpdatingSuspendedHandler(action);
                return;
            }

            action?.Invoke();
        }

        internal IDisposable SuppressTaskPaneRefresh(string reason)
        {
            if (SuppressTaskPaneRefreshHandler != null)
            {
                return SuppressTaskPaneRefreshHandler(reason);
            }

            return new NoOpDisposable();
        }

        internal void RefreshActiveTaskPane(string reason)
        {
            RefreshActiveTaskPaneHandler?.Invoke(reason);
        }

        internal void RequestTaskPaneDisplayForTargetWindow(
            CaseInfoSystem.ExcelAddIn.App.TaskPaneDisplayRequest request,
            Microsoft.Office.Interop.Excel.Workbook workbook,
            Microsoft.Office.Interop.Excel.Window targetWindow)
        {
            RequestTaskPaneDisplayForTargetWindowHandler?.Invoke(request, workbook, targetWindow);
        }

        internal bool ShowKernelSheetAndRefreshPane(string kernelTransitionSheetCodeName, string kernelTransitionReason)
        {
            return ShowKernelSheetAndRefreshPaneHandler == null
                || ShowKernelSheetAndRefreshPaneHandler(kernelTransitionSheetCodeName, kernelTransitionReason);
        }

        internal void ShowWorkbookTaskPaneWhenReady(Microsoft.Office.Interop.Excel.Workbook workbook, string reason)
        {
            ShowWorkbookTaskPaneWhenReadyHandler?.Invoke(workbook, reason);
        }

        private sealed class NoOpDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
