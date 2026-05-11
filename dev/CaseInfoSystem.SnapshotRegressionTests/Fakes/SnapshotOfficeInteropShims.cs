using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Microsoft.Office.Core
{
    public sealed class DocumentProperty
    {
        public string Name { get; set; } = string.Empty;

        public object Value { get; set; }
    }

    public sealed class DocumentProperties : IEnumerable<DocumentProperty>
    {
        private readonly Dictionary<string, DocumentProperty> _properties =
            new Dictionary<string, DocumentProperty>(StringComparer.OrdinalIgnoreCase);

        public DocumentProperty this[string name]
        {
            get
            {
                _properties.TryGetValue(name ?? string.Empty, out DocumentProperty property);
                return property;
            }
        }

        public void Add(string name, bool linkToContent, int type, object value)
        {
            string safeName = name ?? string.Empty;
            _properties[safeName] = new DocumentProperty
            {
                Name = safeName,
                Value = value
            };
        }

        public IEnumerator<DocumentProperty> GetEnumerator()
        {
            return _properties.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}

namespace Microsoft.Office.Interop.Excel
{
    public enum XlWindowState
    {
        xlNormal = -4143,
        xlMinimized = -4140,
        xlMaximized = -4137
    }

    public class Application
    {
        public int Hwnd { get; set; }

        public bool Visible { get; set; } = true;

        public bool ScreenUpdating { get; set; } = true;

        public bool DisplayAlerts { get; set; } = true;

        public bool EnableEvents { get; set; } = true;

        public Workbook ActiveWorkbook { get; set; }

        public Window ActiveWindow { get; set; }

        public Workbooks Workbooks { get; }

        public Application()
        {
            Workbooks = new Workbooks(this);
        }
    }

    public sealed class Workbooks : IEnumerable<Workbook>
    {
        private readonly List<Workbook> _items = new List<Workbook>();

        public Application Owner { get; }

        public Workbooks(Application owner)
        {
            Owner = owner;
        }

        public IEnumerator<Workbook> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(Workbook workbook)
        {
            if (workbook == null || _items.Contains(workbook))
            {
                return;
            }

            workbook.Application = Owner;
            if (workbook.Windows.Count == 0)
            {
                workbook.Windows.Add(new Window());
            }

            _items.Add(workbook);
        }

        public void Remove(Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            _items.Remove(workbook);
        }

        public Func<string, object, bool, Workbook> OpenBehavior { get; set; }

        public Workbook Open(
            string filename,
            object UpdateLinks = null,
            bool ReadOnly = false,
            object Format = null,
            object Password = null,
            object WriteResPassword = null,
            object IgnoreReadOnlyRecommended = null,
            object Origin = null,
            object Delimiter = null,
            object Editable = null,
            object Notify = null,
            object Converter = null,
            object AddToMru = null,
            object Local = null,
            object CorruptLoad = null)
        {
            Workbook workbook = OpenBehavior != null
                ? OpenBehavior(filename, UpdateLinks, ReadOnly)
                : new Workbook
                {
                    FullName = filename ?? string.Empty,
                    Name = Path.GetFileName(filename ?? string.Empty) ?? string.Empty,
                    Path = Path.GetDirectoryName(filename ?? string.Empty) ?? string.Empty
                };

            Add(workbook);
            return workbook;
        }
    }

    public sealed class Workbook
    {
        public Application Application { get; set; }

        public string FullName { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public string Path { get; set; } = string.Empty;

        public Windows Windows { get; } = new Windows();

        public Worksheets Worksheets { get; } = new Worksheets();

        public object CustomDocumentProperties { get; set; }

        public void Close(object saveChanges = null, object filename = null, object routeWorkbook = null)
        {
            Application?.Workbooks.Remove(this);
            if (Application?.ActiveWorkbook == this)
            {
                Application.ActiveWorkbook = null;
                Application.ActiveWindow = null;
            }
        }
    }

    public sealed class Windows : IEnumerable<Window>
    {
        private readonly List<Window> _items = new List<Window>();

        public int Count => _items.Count;

        public Window this[int index]
        {
            get => _items[index - 1];
            set => _items[index - 1] = value;
        }

        public IEnumerator<Window> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(Window window)
        {
            _items.Add(window ?? new Window());
        }
    }

    public sealed class Window
    {
        public int Hwnd { get; set; }

        public bool Visible { get; set; } = true;

        public XlWindowState WindowState { get; set; } = XlWindowState.xlNormal;

        public void Activate()
        {
        }
    }

    public sealed class Worksheets : IEnumerable<Worksheet>
    {
        private readonly List<Worksheet> _items = new List<Worksheet>();

        public Worksheet this[string name] =>
            _items.FirstOrDefault(item => string.Equals(item?.Name, name, StringComparison.OrdinalIgnoreCase));

        public IEnumerator<Worksheet> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(Worksheet worksheet)
        {
            if (worksheet != null)
            {
                _items.Add(worksheet);
            }
        }
    }

    public sealed class Worksheet
    {
        public string CodeName { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public WorksheetCellCollection Cells { get; }

        public WorksheetRows Rows { get; } = new WorksheetRows();

        public WorksheetRangeAccessor Range { get; }

        public Worksheet()
        {
            Cells = new WorksheetCellCollection(this);
            Range = new WorksheetRangeAccessor(this);
        }

        internal Range ResolveCell(int rowIndex, string columnName)
        {
            return Cells[rowIndex, columnName];
        }
    }

    public sealed class WorksheetRows
    {
        public int Count { get; set; } = 1048576;
    }

    public sealed class WorksheetRangeAccessor
    {
        private readonly Worksheet _worksheet;

        internal WorksheetRangeAccessor(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public Range this[object start, object end]
        {
            get
            {
                ParseAddress(Convert.ToString(start) ?? string.Empty, out int startRow, out int startColumn);
                ParseAddress(Convert.ToString(end) ?? string.Empty, out int endRow, out int endColumn);

                int rowCount = Math.Max(0, endRow - startRow + 1);
                int columnCount = Math.Max(0, endColumn - startColumn + 1);
                Array values = Array.CreateInstance(typeof(object), new[] { rowCount, columnCount }, new[] { 1, 1 });
                for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                    {
                        string columnName = ColumnIndexToName(startColumn + columnOffset);
                        Range cell = _worksheet.ResolveCell(startRow + rowOffset, columnName);
                        values.SetValue(cell?.Value2, rowOffset + 1, columnOffset + 1);
                    }
                }

                return new Range(_worksheet)
                {
                    Value2 = values
                };
            }
        }

        private static void ParseAddress(string address, out int rowIndex, out int columnIndex)
        {
            string trimmed = (address ?? string.Empty).Trim().ToUpperInvariant();
            int separator = 0;
            while (separator < trimmed.Length && char.IsLetter(trimmed[separator]))
            {
                separator++;
            }

            string columnName = trimmed.Substring(0, separator);
            string rowPart = trimmed.Substring(separator);
            columnIndex = ColumnNameToIndex(columnName);
            rowIndex = int.TryParse(rowPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedRow)
                ? parsedRow
                : 0;
        }

        private static int ColumnNameToIndex(string columnName)
        {
            int index = 0;
            foreach (char ch in columnName ?? string.Empty)
            {
                index = index * 26 + (ch - 'A' + 1);
            }

            return index;
        }

        private static string ColumnIndexToName(int columnIndex)
        {
            if (columnIndex <= 0)
            {
                return "A";
            }

            string result = string.Empty;
            int remaining = columnIndex;
            while (remaining > 0)
            {
                remaining--;
                result = (char)('A' + (remaining % 26)) + result;
                remaining /= 26;
            }

            return result;
        }
    }

    public sealed class WorksheetCellCollection
    {
        private readonly Worksheet _worksheet;
        private readonly Dictionary<string, Range> _cells = new Dictionary<string, Range>(StringComparer.OrdinalIgnoreCase);

        internal WorksheetCellCollection(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public Range this[object row, object column]
        {
            get
            {
                int rowIndex = Convert.ToInt32(row, CultureInfo.InvariantCulture);
                string columnName = Convert.ToString(column, CultureInfo.InvariantCulture) ?? string.Empty;
                string key = rowIndex.ToString(CultureInfo.InvariantCulture) + "|" + columnName;
                if (!_cells.TryGetValue(key, out Range range))
                {
                    range = new Range(_worksheet)
                    {
                        Row = rowIndex,
                        ColumnName = columnName
                    };
                    _cells[key] = range;
                }

                return range;
            }
        }
    }

    public sealed class Range
    {
        private readonly Worksheet _worksheet;

        internal Range(Worksheet worksheet)
        {
            _worksheet = worksheet;
            Interior = new RangeInterior();
            End = new RangeEndAccessor(this);
        }

        public object Value2 { get; set; }

        public int Row { get; set; }

        internal string ColumnName { get; set; } = string.Empty;

        public RangeInterior Interior { get; }

        public RangeEndAccessor End { get; }

        internal Worksheet Worksheet => _worksheet;
    }

    public sealed class RangeInterior
    {
        public long Color { get; set; }
    }

    public sealed class RangeEndAccessor
    {
        private readonly Range _range;

        internal RangeEndAccessor(Range range)
        {
            _range = range;
        }

        public Range this[object direction]
        {
            get
            {
                int currentRow = _range.Row;
                string columnName = _range.ColumnName;
                for (int rowIndex = currentRow; rowIndex >= 1; rowIndex--)
                {
                    Range candidate = _range.Worksheet.Cells[rowIndex, columnName];
                    string value = Convert.ToString(candidate?.Value2, CultureInfo.InvariantCulture) ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        return candidate;
                    }
                }

                return _range.Worksheet.Cells[1, columnName];
            }
        }
    }
}
