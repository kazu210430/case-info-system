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

    public enum XlCalculation
    {
        xlCalculationAutomatic = -4105,
        xlCalculationManual = -4135
    }

    public enum XlDirection
    {
        xlUp = -4162,
        xlToLeft = -4159
    }

    public enum XlEnableSelection
    {
        xlNoRestrictions = 0,
        xlUnlockedCells = 1,
        xlNoSelection = -4142
    }

    public class Application
    {
        public static List<Application> CreatedApplications { get; } = new List<Application>();

        public static Action<Application> ConfigureNewApplication { get; set; }

        public bool DisplayAlerts { get; set; }

        public bool EnableEvents { get; set; }

        public bool ScreenUpdating { get; set; } = true;

        public XlCalculation Calculation { get; set; } = XlCalculation.xlCalculationAutomatic;

        public bool Ready { get; set; } = true;

        public bool UserControl { get; set; }

        public bool Visible { get; set; }

        public object StatusBar { get; set; }

        public int Hwnd { get; set; }

        public Workbook ActiveWorkbook { get; set; }

        public Window ActiveWindow { get; set; }

        public Workbooks Workbooks { get; }

        public int QuitCallCount { get; private set; }

        public Action QuitBehavior { get; set; }

        public Application()
        {
            Workbooks = new Workbooks(this);
            CreatedApplications.Add(this);
            ConfigureNewApplication?.Invoke(this);
        }

        public void Quit()
        {
            QuitCallCount++;
            QuitBehavior?.Invoke();
        }

        public static void ResetCreatedApplications()
        {
            CreatedApplications.Clear();
            ConfigureNewApplication = null;
        }
    }

    public class Workbooks : List<Workbook>
    {
        public Workbooks()
        {
        }

        public Workbooks(Application owner)
        {
            Owner = owner;
        }

        public Application Owner { get; set; }

        public Func<string, object, bool, Workbook> OpenBehavior { get; set; }

        public new void Add(Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            workbook.Application = Owner;
            if (workbook.Windows.Count == 0)
            {
                workbook.Windows.Add(new Window());
            }

            if (!base.Contains(workbook))
            {
                base.Add(workbook);
            }
        }

        public Workbook Open(string filename, object UpdateLinks = null, bool ReadOnly = false)
        {
            var workbook = OpenBehavior != null
                ? OpenBehavior(filename, UpdateLinks, ReadOnly)
                : new Workbook
                {
                    FullName = filename ?? string.Empty,
                    Name = Path.GetFileName(filename ?? string.Empty),
                    Path = Path.GetDirectoryName(filename ?? string.Empty) ?? string.Empty
                };
            workbook.FullName = filename ?? workbook.FullName ?? string.Empty;
            workbook.Name = string.IsNullOrWhiteSpace(workbook.Name) ? Path.GetFileName(filename ?? string.Empty) : workbook.Name;
            workbook.Path = string.IsNullOrWhiteSpace(workbook.Path) ? (Path.GetDirectoryName(filename ?? string.Empty) ?? string.Empty) : workbook.Path;
            workbook.Application = Owner;
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

        public Windows Windows { get; } = new Windows();

        public Worksheet ActiveSheet { get; set; }

        public Worksheets Worksheets { get; } = new Worksheets();

        public object CustomDocumentProperties { get; set; }

        public int SaveCallCount { get; private set; }

        public Action SaveBehavior { get; set; }

        public int CloseCallCount { get; private set; }

        public Action CloseBehavior { get; set; }

        public bool? LastCloseSaveChanges { get; private set; }

        public object LastCloseSaveChangesArgument { get; private set; }

        public object LastCloseFilename { get; private set; }

        public object LastCloseRouteWorkbook { get; private set; }

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

        public void Close(object SaveChanges = null, object Filename = null, object RouteWorkbook = null)
        {
            CloseCallCount++;
            LastCloseSaveChangesArgument = SaveChanges;
            LastCloseSaveChanges = SaveChanges is bool saveChangesValue
                ? saveChangesValue
                : (bool?)null;
            LastCloseFilename = Filename;
            LastCloseRouteWorkbook = RouteWorkbook;
            CloseBehavior?.Invoke();
            Application?.Workbooks?.Remove(this);
            if (Application?.ActiveWorkbook == this)
            {
                Application.ActiveWorkbook = null;
                Application.ActiveWindow = null;
            }
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

    public class Windows : WorkbookWindows
    {
    }

    public class Window
    {
        public string Caption { get; set; } = string.Empty;

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

        public bool ProtectContents { get; set; }

        public bool ProtectDrawingObjects { get; set; }

        public bool ProtectScenarios { get; set; }

        public XlEnableSelection EnableSelection { get; set; }

        public WorksheetCellCollection Cells { get; } = new WorksheetCellCollection();

        public WorksheetRowCollection Rows { get; } = new WorksheetRowCollection();

        public WorksheetColumnCollection Columns { get; } = new WorksheetColumnCollection();

        public WorksheetRangeAccessor Range { get; } = new WorksheetRangeAccessor();

        public void Activate()
        {
        }

        public void Unprotect(string Password = null)
        {
            ProtectContents = false;
            ProtectDrawingObjects = false;
            ProtectScenarios = false;
        }

        public void Protect(string Password = null, bool UserInterfaceOnly = false, bool AllowFiltering = false, bool AllowSorting = false)
        {
            ProtectContents = true;
            ProtectDrawingObjects = true;
            ProtectScenarios = true;
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
        public Range()
        {
            End = new RangeEndAccessor(this);
        }

        public object Value2 { get; set; }

        public int Row { get; set; }

        public bool Locked { get; set; } = true;

        public RangeEndAccessor End { get; }

        public string Text
        {
            get => Convert.ToString(Value2) ?? string.Empty;
            set => Value2 = value;
        }
    }

    public sealed class RangeEndAccessor
    {
        private readonly Range _owner;

        public RangeEndAccessor(Range owner)
        {
            _owner = owner;
        }

        public Range this[XlDirection direction] => _owner;
    }

    public sealed class WorksheetRowCollection
    {
        public int Count { get; set; } = 1048576;
    }

    public sealed class WorksheetColumnCollection
    {
        public int Count { get; set; } = 16384;

        public Range this[object column] => new Range();
    }

    public sealed class WorksheetRangeAccessor
    {
        public Range this[object from, object to] => new Range();
    }

    public sealed class WorksheetCellCollection : Range
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
                    if (int.TryParse(Convert.ToString(row), out int rowNumber))
                    {
                        range.Row = rowNumber;
                    }
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

        internal Func<CaseInfoSystem.ExcelAddIn.Domain.WorkbookContext, string, string, bool> ShowKernelSheetAndRefreshPaneFromHomeHandler { get; set; }

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

        internal bool ShowKernelSheetAndRefreshPaneFromHome(
            CaseInfoSystem.ExcelAddIn.Domain.WorkbookContext context,
            string kernelTransitionSheetCodeName,
            string kernelTransitionReason,
            out Microsoft.Office.Interop.Excel.Workbook displayedWorkbook)
        {
            displayedWorkbook = context == null ? null : context.Workbook;
            return ShowKernelSheetAndRefreshPaneFromHomeHandler != null
                ? ShowKernelSheetAndRefreshPaneFromHomeHandler(context, kernelTransitionSheetCodeName, kernelTransitionReason)
                : true;
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
