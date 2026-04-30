using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class CaseTaskPaneViewState
    {
        private readonly List<CaseTaskPaneActionViewState> _specialButtons;
        private readonly List<CaseTaskPaneTabPageViewState> _tabPages;

        internal CaseTaskPaneViewState(
            string statusMessage,
            string selectedTabName,
            IEnumerable<CaseTaskPaneActionViewState> specialButtons,
            IEnumerable<CaseTaskPaneTabPageViewState> tabPages)
        {
            StatusMessage = statusMessage ?? string.Empty;
            SelectedTabName = selectedTabName ?? string.Empty;
            _specialButtons = specialButtons == null ? new List<CaseTaskPaneActionViewState>() : new List<CaseTaskPaneActionViewState>(specialButtons);
            _tabPages = tabPages == null ? new List<CaseTaskPaneTabPageViewState>() : new List<CaseTaskPaneTabPageViewState>(tabPages);
        }

        internal string StatusMessage { get; }

        internal string SelectedTabName { get; }

        internal IReadOnlyList<CaseTaskPaneActionViewState> SpecialButtons
        {
            get { return _specialButtons; }
        }

        internal IReadOnlyList<CaseTaskPaneTabPageViewState> TabPages
        {
            get { return _tabPages; }
        }

        internal bool HasStatusMessage
        {
            get { return !string.IsNullOrWhiteSpace(StatusMessage); }
        }

        internal CaseTaskPaneViewState WithSelectedTab(string tabName)
        {
            if (HasStatusMessage || _tabPages.Count == 0)
            {
                return this;
            }

            string resolvedTabName = ResolveSelectedTabName(tabName);
            if (string.Equals(resolvedTabName, SelectedTabName, StringComparison.Ordinal))
            {
                return this;
            }

            return new CaseTaskPaneViewState(StatusMessage, resolvedTabName, _specialButtons, _tabPages);
        }

        internal CaseTaskPaneTabPageViewState GetSelectedTabPage()
        {
            if (_tabPages.Count == 0)
            {
                return null;
            }

            string resolvedTabName = ResolveSelectedTabName(SelectedTabName);
            foreach (CaseTaskPaneTabPageViewState tabPage in _tabPages)
            {
                if (string.Equals(tabPage.TabName, resolvedTabName, StringComparison.Ordinal))
                {
                    return tabPage;
                }
            }

            return _tabPages[0];
        }

        private string ResolveSelectedTabName(string tabName)
        {
            if (_tabPages.Count == 0)
            {
                return string.Empty;
            }

            foreach (CaseTaskPaneTabPageViewState tabPage in _tabPages)
            {
                if (string.Equals(tabPage.TabName, tabName, StringComparison.Ordinal))
                {
                    return tabPage.TabName;
                }
            }

            return _tabPages[0].TabName;
        }
    }

    internal sealed class CaseTaskPaneActionViewState
    {
        internal string Caption { get; set; }
        internal string ActionKind { get; set; }
        internal string Key { get; set; }
        internal Color BackColor { get; set; }
    }

    internal sealed class CaseTaskPaneTabPageViewState
    {
        private readonly List<CaseTaskPaneActionViewState> _documentButtons = new List<CaseTaskPaneActionViewState>();

        internal int Order { get; set; }
        internal string TabName { get; set; }
        internal Color BackColor { get; set; }

        internal IReadOnlyList<CaseTaskPaneActionViewState> DocumentButtons
        {
            get { return _documentButtons; }
        }

        internal void SetDocumentButtons(IEnumerable<CaseTaskPaneActionViewState> documentButtons)
        {
            _documentButtons.Clear();
            if (documentButtons != null)
            {
                _documentButtons.AddRange(documentButtons);
            }
        }
    }

    internal sealed class CaseTaskPaneViewStateBuilder
    {
        private const string AllTabName = "\u5168\u3066";

        /// <summary>
        /// CASE Task Pane の描画用 view state を組み立てる。
        /// </summary>
        internal CaseTaskPaneViewState Build(TaskPaneSnapshot snapshot, string selectedTabName)
        {
            if (snapshot == null)
            {
                return CreateMessageState("Failed to load definitions.");
            }

            if (snapshot.HasError)
            {
                return CreateMessageState(snapshot.ErrorMessage);
            }

            IReadOnlyList<TaskPaneActionDefinition> specialButtons = snapshot.SpecialButtons;
            IReadOnlyList<TaskPaneTabDefinition> tabs = snapshot.Tabs;
            IReadOnlyList<TaskPaneDocDefinition> documents = snapshot.DocButtons;

            List<CaseTaskPaneTabPageViewState> tabPages = BuildTabPages(tabs, documents);
            if (tabPages.Count == 0)
            {
                return CreateMessageState("No available document buttons.");
            }

            string resolvedTabName = ResolveSelectedTabName(tabPages, selectedTabName);
            return new CaseTaskPaneViewState(string.Empty, resolvedTabName, BuildSpecialButtons(specialButtons), tabPages);
        }

        internal CaseTaskPaneViewState BuildWorkbookNotFoundState()
        {
            return CreateMessageState("\u5BFE\u8C61\u306E CASE \u30D6\u30C3\u30AF\u304C\u898B\u3064\u304B\u308A\u307E\u305B\u3093\u3002");
        }

        internal CaseTaskPaneViewState BuildActionFailedState()
        {
            return CreateMessageState("\u64CD\u4F5C\u306E\u5B9F\u884C\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002");
        }

        private static CaseTaskPaneViewState CreateMessageState(string message)
        {
            return new CaseTaskPaneViewState(message ?? string.Empty, string.Empty, Array.Empty<CaseTaskPaneActionViewState>(), Array.Empty<CaseTaskPaneTabPageViewState>());
        }

        private static List<CaseTaskPaneActionViewState> BuildSpecialButtons(IReadOnlyList<TaskPaneActionDefinition> specialButtons)
        {
            var buttons = new List<CaseTaskPaneActionViewState>();
            foreach (TaskPaneActionDefinition action in specialButtons ?? Array.Empty<TaskPaneActionDefinition>())
            {
                if (action == null)
                {
                    continue;
                }

                buttons.Add(new CaseTaskPaneActionViewState
                {
                    Caption = action.Caption ?? string.Empty,
                    ActionKind = action.ActionKind ?? string.Empty,
                    Key = action.Key ?? string.Empty,
                    BackColor = action.BackColor
                });
            }

            return buttons;
        }

        private static List<CaseTaskPaneTabPageViewState> BuildTabPages(IReadOnlyList<TaskPaneTabDefinition> tabs, IReadOnlyList<TaskPaneDocDefinition> documents)
        {
            IEnumerable<TaskPaneTabDefinition> sourceTabs = tabs ?? Array.Empty<TaskPaneTabDefinition>();
            var orderedTabs = sourceTabs.Where(tab => tab != null).OrderBy(tab => tab.Order).ToList();
            var tabPages = new List<CaseTaskPaneTabPageViewState>(orderedTabs.Count);
            foreach (TaskPaneTabDefinition tab in orderedTabs)
            {
                var page = new CaseTaskPaneTabPageViewState
                {
                    Order = tab.Order,
                    TabName = tab.TabName ?? string.Empty,
                    BackColor = tab.BackColor
                };
                page.SetDocumentButtons(BuildDocumentButtonsForTab(documents, page.TabName));
                tabPages.Add(page);
            }

            return tabPages;
        }

        private static IEnumerable<CaseTaskPaneActionViewState> BuildDocumentButtonsForTab(IReadOnlyList<TaskPaneDocDefinition> documents, string tabName)
        {
            IEnumerable<TaskPaneDocDefinition> source = documents ?? Array.Empty<TaskPaneDocDefinition>();
            // Preserve the previous UI-side tab filtering and ordering rules in the app layer.
            IEnumerable<TaskPaneDocDefinition> filteredDocuments = IsAllTab(tabName)
                ? source.Where(document => document != null).OrderBy(document => ParseDocOrderKey(document.Key)).ThenBy(document => document.Key ?? string.Empty, StringComparer.Ordinal)
                : source.Where(document => document != null && string.Equals(document.TabName, tabName, StringComparison.Ordinal)).OrderBy(document => document.RowIndex).ThenBy(document => document.Key ?? string.Empty, StringComparer.Ordinal);

            foreach (TaskPaneDocDefinition document in filteredDocuments)
            {
                yield return new CaseTaskPaneActionViewState
                {
                    Caption = document.Caption ?? string.Empty,
                    ActionKind = document.ActionKind ?? string.Empty,
                    Key = document.Key ?? string.Empty,
                    BackColor = document.FillColor
                };
            }
        }

        private static string ResolveSelectedTabName(IReadOnlyList<CaseTaskPaneTabPageViewState> tabPages, string selectedTabName)
        {
            if (tabPages == null || tabPages.Count == 0)
            {
                return string.Empty;
            }

            foreach (CaseTaskPaneTabPageViewState tabPage in tabPages)
            {
                if (string.Equals(tabPage.TabName, selectedTabName, StringComparison.Ordinal))
                {
                    return tabPage.TabName;
                }
            }

            return tabPages[0].TabName;
        }

        private static bool IsAllTab(string tabName)
        {
            return string.Equals(tabName, AllTabName, StringComparison.Ordinal);
        }

        private static int ParseDocOrderKey(string key)
        {
            return int.TryParse(key, out int result) ? result : int.MaxValue;
        }
    }
}
