using System;
using System.Collections;
using System.Collections.Generic;

namespace Microsoft.Office.Interop.Word
{
    public enum WdContentControlType
    {
        wdContentControlRichText = 0,
        wdContentControlText = 1,
        wdContentControlDate = 6,
        wdContentControlCheckBox = 8
    }

    public sealed class Document
    {
        public ContentControls ContentControls { get; } = new ContentControls();
    }

    public sealed class Selection
    {
        public Range Range { get; } = new Range();

        public void SetRange(int start, int end)
        {
            Range.Start = start;
            Range.End = end;
        }
    }

    public sealed class ContentControls : IEnumerable
    {
        private readonly List<ContentControl> _controls = new List<ContentControl>();

        public int Count => _controls.Count;

        public ContentControl Item(int index)
        {
            return _controls[index - 1];
        }

        public void Add(ContentControl control)
        {
            _controls.Add(control);
        }

        public IEnumerator GetEnumerator()
        {
            return _controls.GetEnumerator();
        }
    }

    public sealed class ContentControl
    {
        private string _tag = string.Empty;

        private string _title = string.Empty;

        public string Tag
        {
            get => _tag;
            set
            {
                if (ThrowOnTagSet)
                {
                    throw new InvalidOperationException("tag write failed");
                }

                _tag = value ?? string.Empty;
            }
        }

        public string Title
        {
            get => _title;
            set
            {
                if (ThrowOnTitleSet)
                {
                    throw new InvalidOperationException("title write failed");
                }

                _title = value ?? string.Empty;
            }
        }

        public WdContentControlType Type { get; set; } = WdContentControlType.wdContentControlText;

        public Range Range { get; } = new Range();

        public bool ThrowOnTagSet { get; set; }

        public bool ThrowOnTitleSet { get; set; }
    }

    public sealed class Range
    {
        public int Start { get; set; }

        public int End { get; set; }

        public string Text { get; set; } = string.Empty;
    }
}
