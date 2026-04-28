using System;
using System.Drawing;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class TaskPaneSnapshotParser
    {
        internal static TaskPaneSnapshot Parse(string snapshotText)
        {
            TaskPaneSnapshot snapshot = new TaskPaneSnapshot();

            if (string.IsNullOrWhiteSpace(snapshotText))
            {
                snapshot.HasError = true;
                snapshot.ErrorMessage = "Task Pane snapshot could not be loaded.";
                return snapshot;
            }

            string[] lines = snapshotText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawLine in lines)
            {
                string[] fields = SplitFields(rawLine);
                if (fields.Length == 0)
                {
                    continue;
                }

                switch (fields[0])
                {
                    case "META":
                        ParseMeta(snapshot, fields);
                        break;
                    case "SPECIAL":
                        ParseSpecial(snapshot, fields);
                        break;
                    case "TAB":
                        ParseTab(snapshot, fields);
                        break;
                    case "DOC":
                        ParseDoc(snapshot, fields);
                        break;
                }
            }

            if (string.Equals(snapshot.WorkbookName, "ERROR", StringComparison.OrdinalIgnoreCase))
            {
                snapshot.HasError = true;
                snapshot.ErrorMessage = snapshot.WorkbookPath;
            }

            return snapshot;
        }

        private static void ParseMeta(TaskPaneSnapshot snapshot, string[] fields)
        {
            if (fields.Length > 2)
            {
                snapshot.WorkbookName = fields[2];
            }

            if (fields.Length > 3)
            {
                snapshot.WorkbookPath = fields[3];
            }

            if (fields.Length > 4)
            {
                snapshot.PreferredPaneWidth = ParseInt(GetField(fields, 4));
            }

        }

        private static void ParseSpecial(TaskPaneSnapshot snapshot, string[] fields)
        {
            TaskPaneActionDefinition action = new TaskPaneActionDefinition
            {
                Name = GetField(fields, 1),
                Caption = GetField(fields, 2),
                ActionKind = GetField(fields, 3),
                Key = GetField(fields, 4),
                Left = ParseInt(GetField(fields, 5)),
                Top = ParseInt(GetField(fields, 6)),
                Width = ParseInt(GetField(fields, 7)),
                Height = ParseInt(GetField(fields, 8)),
                BackColor = ParseColor(GetField(fields, 9))
            };

            snapshot.SpecialButtons.Add(action);
        }

        private static void ParseTab(TaskPaneSnapshot snapshot, string[] fields)
        {
            TaskPaneTabDefinition tab = new TaskPaneTabDefinition
            {
                Order = ParseInt(GetField(fields, 1)),
                TabName = GetField(fields, 2),
                BackColor = ParseColor(GetField(fields, 3))
            };

            snapshot.Tabs.Add(tab);
        }

        private static void ParseDoc(TaskPaneSnapshot snapshot, string[] fields)
        {
            TaskPaneDocDefinition doc = new TaskPaneDocDefinition
            {
                Name = GetField(fields, 1),
                Key = GetField(fields, 2),
                Caption = GetField(fields, 3),
                ActionKind = GetField(fields, 4),
                TabName = GetField(fields, 5),
                RowIndex = ParseInt(GetField(fields, 6)),
                FillColor = ParseColor(GetField(fields, 7)),
                TemplateFileName = GetField(fields, 8)
            };

            snapshot.DocButtons.Add(doc);
        }

        private static string[] SplitFields(string line)
        {
            string[] rawFields = line.Split('\t');
            for (int i = 0; i < rawFields.Length; i++)
            {
                rawFields[i] = Unescape(rawFields[i]);
            }

            return rawFields;
        }

        private static string Unescape(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            return value
                .Replace("\\n", "\n")
                .Replace("\\t", "\t")
                .Replace("\\\\", "\\");
        }

        private static string GetField(string[] fields, int index)
        {
            return index < fields.Length ? fields[index] : string.Empty;
        }

        private static int ParseInt(string value)
        {
            int result;
            return int.TryParse(value, out result) ? result : 0;
        }

        private static Color ParseColor(string value)
        {
            int oleColor;
            if (!int.TryParse(value, out oleColor) || oleColor == 0)
            {
                return Color.Empty;
            }

            return ColorTranslator.FromOle(oleColor);
        }
    }
}
